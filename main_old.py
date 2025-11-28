# main.py
import sys
import time
import traceback
import os

# 瀵煎叆鎵€鏈夋ā鍧?
from config import *
from etabs_api_loader import load_dotnet_etabs_api
from etabs_setup import setup_etabs
from materials_sections import define_all_materials_and_sections
from response_spectrum import define_response_spectrum_functions_in_etabs
from load_cases import define_all_load_cases
from frame_geometry import create_frame_structure
from load_assignment import assign_all_loads_to_frame_structure
from analysis_module import wait_and_run_analysis, check_analysis_completion
from results_extraction import extract_all_analysis_results
from file_operations import finalize_and_save_model, cleanup_etabs_on_error, check_output_directory
from results_extraction.member_forces import extract_and_save_frame_forces

# 灏濊瘯浠?design_module 瀵煎叆涓诲嚱鏁?
try:
    from design_module import perform_concrete_design_and_extract_results

    design_module_available = True
    print("鉁?璁捐妯″潡瀵煎叆鎴愬姛")
except ImportError as e:
    design_module_available = False
    print(f"鈿狅笍 瀵煎叆璁捐妯″潡鏃跺嚭鐜伴棶棰? {e}")
    print("灏嗚烦杩囪璁″姛鑳?..")


    # 瀹氫箟涓€涓┖鐨勬浛浠ｅ嚱鏁帮紝浣垮叾鍦ㄦ湭瀵煎叆鏃朵篃鑳芥甯歌皟鐢?
    def perform_concrete_design_and_extract_results():
        print("鈴笍 璁捐妯″潡瀵煎叆澶辫触锛岃烦杩囨瀯浠惰璁°€?)
        return False  # 杩斿洖 False 琛ㄧず澶辫触

# 尝试导入设计内力提取模块（优先新的 results_extraction 入口）
design_force_extraction_available = False
extract_design_forces_and_summary = None

try:
    from results_extraction.design_forces import extract_design_forces_and_summary

    design_force_extraction_available = True
    print("✅设计内力提取模块导入成功: results_extraction.design_forces")
except Exception as primary_error:
    print(f"⚠️ 首选设计内力提取模块导入失败: {primary_error}")

    possible_modules = [
        'design_force_extraction',
        'design_force_extraction_fixed',
        'design_force_extraction_improved',
    ]

    for module_name in possible_modules:
        try:
            module = __import__(module_name, fromlist=['extract_design_forces_and_summary'])
            extract_design_forces_and_summary = getattr(module, 'extract_design_forces_and_summary')
            design_force_extraction_available = True
            print(f"✅设计内力提取模块导入成功: {module_name}")
            break
        except ImportError as e:
            print(f"⚠️ 尝试导入 {module_name} 失败: {e}")
            continue

if not design_force_extraction_available:
    print("⚠️ 所有设计内力提取模块导入失败，将跳过设计内力提取功能。")

    def extract_design_forces_and_summary(column_names, beam_names):
        print("⏭️ 设计内力提取模块导入失败，跳过设计内力提取。")
        return False


def print_project_info():
    """鎵撳嵃椤圭洰淇℃伅"""
    print("=" * 80)
    print("ETABS 妗嗘灦缁撴瀯鑷姩寤烘ā鑴氭湰 v6.3.1 (璁捐妯″潡 v12.1)")
    print("=" * 80)
    print("椤圭洰鐗圭偣锛?)
    print("1. 10灞傞挗绛嬫贩鍑濆湡妗嗘灦缁撴瀯")
    print("2. 閲囩敤妗嗘灦鏌卞拰妗嗘灦姊佷綋绯?)
    print("3. 妤兼澘璁剧疆涓鸿啘鍗曞厓锛堥潰澶栧垰搴︿负0锛?)
    print("4. 鍩轰簬GB50011-2010鍙嶅簲璋卞垎鏋?)
    print("5. 鑷姩鎻愬彇妯℃€佷俊鎭€佸眰闂翠綅绉昏鍜屾瀯浠跺唴鍔?)
    print("6. 鎵цGB50010-2010娣峰嚌鍦熸瀯浠堕厤绛嬭璁?)
    print("7. 鎻愬彇鏋勪欢璁捐鍐呭姏鏁版嵁")
    print("8. 瀹屽叏妯″潡鍖栬璁★紝渚夸簬缁存姢鍜屾墿灞?)
    print()
    print("妯″潡鐘舵€侊細")
    print(f"- 璁捐妯″潡: {'鉁?鍙敤' if design_module_available else '鉂?涓嶅彲鐢?}")
    print(f"- 璁捐鍐呭姏鎻愬彇妯″潡: {'鉁?鍙敤' if design_force_extraction_available else '鉂?涓嶅彲鐢?}")
    print()
    print("缁撴瀯鍙傛暟锛?)
    print(f"- 妤煎眰鏁帮細{NUM_STORIES}灞?)
    print(f"- 缃戞牸锛歿NUM_GRID_LINES_X}脳{NUM_GRID_LINES_Y} ({SPACING_X}m脳{SPACING_Y}m)")
    print(f"- 妗嗘灦鏌憋細{FRAME_COLUMN_WIDTH}m脳{FRAME_COLUMN_HEIGHT}m")
    print(f"- 妗嗘灦姊侊細{FRAME_BEAM_WIDTH}m脳{FRAME_BEAM_HEIGHT}m")
    print(f"- 妤兼澘鍘氬害锛歿SLAB_THICKNESS}m (鑶滃崟鍏?")
    print(f"- 灞傞珮锛氶灞倇BOTTOM_STORY_HEIGHT}m锛屾爣鍑嗗眰{TYPICAL_STORY_HEIGHT}m")
    print(f"- 鎬婚珮搴︼細{BOTTOM_STORY_HEIGHT + (NUM_STORIES - 1) * TYPICAL_STORY_HEIGHT:.1f}m")
    print()
    print("鍦伴渿鍙傛暟锛?)
    print(f"- 璁鹃槻鐑堝害锛歿RS_DESIGN_INTENSITY}搴?)
    print(f"- 鏈€澶у湴闇囧奖鍝嶇郴鏁帮細{RS_BASE_ACCEL_G}")
    print(f"- 鍦哄湴绫诲埆锛歿RS_SITE_CLASS}绫?)
    print(f"- 鐗瑰緛鍛ㄦ湡锛歿RS_CHARACTERISTIC_PERIOD}s")
    print(f"- 鍦伴渿鍒嗙粍锛氱{RS_SEISMIC_GROUP}缁?)
    print()
    print("璁捐鍙傛暟锛?)
    print(f"- 浣跨敤ETABS榛樿娣峰嚌鍦熻璁¤鑼?)
    print(f"- 鏄惁鎵ц閰嶇瓔璁捐锛歿'鏄? if PERFORM_CONCRETE_DESIGN else '鍚?}")
    print(f"- 鏄惁鎻愬彇璁捐鍐呭姏锛歿'鏄? if PERFORM_CONCRETE_DESIGN and design_force_extraction_available else '鍚?}")
    print("=" * 80)


def main():
    """涓诲嚱鏁?- 妗嗘灦缁撴瀯寤烘ā娴佺▼"""
    script_start_time = time.time()

    # 鎵撳嵃椤圭洰淇℃伅
    print_project_info()

    # 鍒濆鍖栧彉閲忥紝浠ラ槻鏌愪簺闃舵琚烦杩?
    column_names, beam_names, slab_names, story_heights = [], [], [], {}

    try:
        # ========== 绗竴闃舵锛氬垵濮嬪寲 ==========
        print("\n馃殌 绗竴闃舵锛氱郴缁熷垵濮嬪寲")
        if not check_output_directory(): sys.exit(1)
        load_dotnet_etabs_api()
        _, sap_model = setup_etabs()

        # ========== 绗簩闃舵锛氭ā鍨嬪畾涔?==========
        print("\n馃彈锔?绗簩闃舵锛氭ā鍨嬪畾涔?)
        define_all_materials_and_sections()
        define_response_spectrum_functions_in_etabs()
        define_all_load_cases()

        # ========== 绗笁闃舵锛氬嚑浣曞缓妯?==========
        print("\n馃彚 绗笁闃舵锛氭鏋剁粨鏋勫缓妯?)
        column_names, beam_names, slab_names, story_heights = create_frame_structure()

        # ========== 绗洓闃舵锛氳嵎杞藉垎閰?==========
        print("\n鈿栵笍 绗洓闃舵锛氳嵎杞藉垎閰?)
        assign_all_loads_to_frame_structure(column_names, beam_names, slab_names)

        # ========== 绗簲闃舵锛氫繚瀛樻ā鍨?==========
        print("\n馃捑 绗簲闃舵锛氫繚瀛樻ā鍨?)
        finalize_and_save_model()

        # ========== 绗叚闃舵锛氱粨鏋勫垎鏋?==========
        print("\n馃攳 绗叚闃舵锛氱粨鏋勫垎鏋?)
        wait_and_run_analysis(5)
        if not check_analysis_completion():
            print("鈿狅笍 鍒嗘瀽鐘舵€佹鏌ュ紓甯革紝浣嗙户缁皾璇曟彁鍙栫粨鏋?)

        # ========== 绗竷闃舵锛氱粨鏋滄彁鍙?==========
        print("\n馃搳 绗竷闃舵锛氱粨鏋滄彁鍙?)
        extract_all_analysis_results()
        extract_and_save_frame_forces(column_names + beam_names)

        # ========== 绗叓闃舵锛氭瀯浠惰璁?==========
        design_completed_successfully = False
        if PERFORM_CONCRETE_DESIGN and design_module_available:
            print("\n馃彈锔?绗叓闃舵锛氭贩鍑濆湡鏋勪欢閰嶇瓔璁捐")
            try:
                # 鍙皟鐢ㄤ富鍑芥暟锛屽畠浼氬鐞嗘墍鏈夊唴閮ㄩ€昏緫鍜岄敊璇?
                design_completed_successfully = perform_concrete_design_and_extract_results()

                if design_completed_successfully:
                    print("鉁?璁捐鍜岀粨鏋滄彁鍙栭獙璇侀€氳繃銆?)
                else:
                    print("鈿狅笍 璁捐鍜岀粨鏋滄彁鍙栨湭鎴愬姛锛岃妫€鏌ヤ互涓?design_module 鏃ュ織銆?)

            except Exception as design_error:
                print(f"鈿狅笍 鏋勪欢璁捐妯″潡鍙戠敓鏈崟鑾风殑涓ラ噸閿欒: {design_error}")
                print("閿欒璇︽儏:")
                traceback.print_exc()

            finally:
                print("鉁?鏋勪欢璁捐闃舵瀹屾垚銆?)  # 鏃犺鎴愬姛涓庡惁閮芥爣璁伴樁娈靛畬鎴?
        elif PERFORM_CONCRETE_DESIGN and not design_module_available:
            print("\n鈴笍 绗叓闃舵锛氳烦杩囨瀯浠惰璁★紙璁捐妯″潡涓嶅彲鐢級銆?)
        else:
            print("\n鈴笍 绗叓闃舵锛氳烦杩囨瀯浠惰璁★紙鐢眂onfig鏂囦欢璁剧疆锛夈€?)

        # ========== 绗節闃舵锛氭瀯浠惰璁″唴鍔涙彁鍙?==========
        design_force_extraction_successful = False
        if (PERFORM_CONCRETE_DESIGN and design_completed_successfully and
                design_force_extraction_available):
            print("\n馃敩 绗節闃舵锛氭瀯浠惰璁″唴鍔涙彁鍙?)
            try:
                print("姝ｅ湪鎻愬彇妗嗘灦鏌卞拰妗嗘灦姊佺殑璁捐鍐呭姏...")
                design_force_extraction_successful = extract_design_forces_and_summary(
                    column_names, beam_names
                )

                if design_force_extraction_successful:
                    print("鉁?鏋勪欢璁捐鍐呭姏鎻愬彇鎴愬姛銆?)
                else:
                    print("鈿狅笍 鏋勪欢璁捐鍐呭姏鎻愬彇澶辫触锛岃妫€鏌ユ棩蹇椼€?)

            except Exception as extraction_error:
                print(f"鈿狅笍 鏋勪欢璁捐鍐呭姏鎻愬彇妯″潡鍙戠敓閿欒: {extraction_error}")
                print("閿欒璇︽儏:")
                traceback.print_exc()

            finally:
                print("鉁?鏋勪欢璁捐鍐呭姏鎻愬彇闃舵瀹屾垚銆?)
        elif PERFORM_CONCRETE_DESIGN and design_completed_successfully and not design_force_extraction_available:
            print("\n鈴笍 绗節闃舵锛氳烦杩囨瀯浠惰璁″唴鍔涙彁鍙栵紙鎻愬彇妯″潡涓嶅彲鐢級銆?)
        elif PERFORM_CONCRETE_DESIGN and not design_completed_successfully:
            print("\n鈴笍 绗節闃舵锛氳烦杩囨瀯浠惰璁″唴鍔涙彁鍙栵紙璁捐闃舵鏈垚鍔熷畬鎴愶級銆?)
        else:
            print("\n鈴笍 绗節闃舵锛氳烦杩囨瀯浠惰璁″唴鍔涙彁鍙栵紙鏈墽琛屾瀯浠惰璁★級銆?)

        # ========== 瀹屾垚 ==========
        elapsed_time = time.time() - script_start_time
        print("\n" + "=" * 80)
        print("馃帀 妗嗘灦缁撴瀯寤烘ā瀹屾垚锛?)
        print("=" * 80)
        print("鉁?涓昏瀹屾垚鍔熻兘:")
        print(f"   1. {NUM_STORIES}灞傞挗绛嬫贩鍑濆湡妗嗘灦缁撴瀯寤烘ā")
        print(f"   2. 鍒涘缓浜?{len(column_names)} 鏍规鏋舵煴")
        print(f"   3. 鍒涘缓浜?{len(beam_names)} 鏍规鏋舵")
        print(f"   4. 鍒涘缓浜?{len(slab_names)} 鍧楁ゼ鏉匡紙鑶滃崟鍏冿級")
        print("   5. 瀹屾垚浜嗚嵎杞藉垎閰嶅拰鍦伴渿鍙傛暟璁剧疆")
        print("   6. 瀹屾垚浜嗘ā鎬佸垎鏋愬拰鍙嶅簲璋卞垎鏋?)
        print("   7. 鎻愬彇浜嗘ā鎬佷俊鎭€佸眰闂翠綅绉昏鍜屾瀯浠跺唴鍔?)
        if PERFORM_CONCRETE_DESIGN and design_module_available:
            if design_completed_successfully:
                print("   8. 鎴愬姛瀹屾垚娣峰嚌鍦熸瀯浠堕厤绛嬭璁″拰缁撴灉鎻愬彇銆?)
                if design_force_extraction_successful:
                    print("   9. 鎴愬姛鎻愬彇鏋勪欢璁捐鍐呭姏鏁版嵁銆?)
                else:
                    print("   9. 鏋勪欢璁捐鍐呭姏鎻愬彇鎵ц瀹屾瘯锛屼絾鏈垚鍔熴€?)
            else:
                print("   8. 娣峰嚌鍦熸瀯浠堕厤绛嬭璁℃墽琛屽畬姣曪紝浣嗙粨鏋滄彁鍙栨垨楠岃瘉澶辫触銆?)
                print("   9. 璺宠繃鏋勪欢璁捐鍐呭姏鎻愬彇銆?)
        else:
            if not design_module_available:
                print("   8. 璁捐妯″潡涓嶅彲鐢紝璺宠繃娣峰嚌鍦熸瀯浠堕厤绛嬭璁°€?)
            print("   9. 璺宠繃鏋勪欢璁捐鍐呭姏鎻愬彇銆?)
        print()
        print("馃搧 杈撳嚭鏂囦欢:")
        print(f"   妯″瀷鏂囦欢: {MODEL_PATH}")
        print(f"   鏋勪欢鍐呭姏: {os.path.join(SCRIPT_DIRECTORY, 'frame_member_forces.csv')}")
        if PERFORM_CONCRETE_DESIGN and design_module_available:
            print(f"   閰嶇瓔璁捐: {os.path.join(SCRIPT_DIRECTORY, 'concrete_design_results.csv')}")
            print(f"   璁捐鎶ュ憡: {os.path.join(SCRIPT_DIRECTORY, 'design_summary_report.txt')}")
            if design_force_extraction_successful:
                print(f"   鏌辫璁″唴鍔? {os.path.join(SCRIPT_DIRECTORY, 'column_design_forces.csv')}")
                print(f"   姊佽璁″唴鍔? {os.path.join(SCRIPT_DIRECTORY, 'beam_design_forces.csv')}")
                print(f"   鍐呭姏姹囨€? {os.path.join(SCRIPT_DIRECTORY, 'design_forces_summary_report.txt')}")
        print()
        print("馃彈锔?缁撴瀯淇℃伅:")
        total_height = BOTTOM_STORY_HEIGHT + (NUM_STORIES - 1) * TYPICAL_STORY_HEIGHT if NUM_STORIES > 0 else 0
        print(f"   缁撴瀯绫诲瀷: {NUM_STORIES}灞傞挗绛嬫贩鍑濆湡妗嗘灦缁撴瀯")
        print(f"   骞抽潰灏哄: {(NUM_GRID_LINES_X - 1) * SPACING_X:.1f}m 脳 {(NUM_GRID_LINES_Y - 1) * SPACING_Y:.1f}m")
        print(f"   缁撴瀯鎬婚珮: {total_height:.1f}m")
        print(f"   鎶楅渿璁鹃槻: {RS_DESIGN_INTENSITY}搴︼紝{RS_SITE_CLASS}绫诲満鍦?)
        print()
        print(f"鈴憋笍 鎬绘墽琛屾椂闂? {elapsed_time:.2f} 绉?)

        # 杈撳嚭鎵ц鐘舵€佹€荤粨
        print("\n馃搵 鎵ц鐘舵€佹€荤粨:")
        print(f"   鉁?缁撴瀯寤烘ā: 鎴愬姛")
        print(f"   鉁?缁撴瀯鍒嗘瀽: 鎴愬姛")
        print(f"   鉁?缁撴灉鎻愬彇: 鎴愬姛")
        if PERFORM_CONCRETE_DESIGN:
            if design_module_available:
                status_design = "鎴愬姛" if design_completed_successfully else "澶辫触"
                print(f"   {'鉁? if design_completed_successfully else '鉂?} 鏋勪欢璁捐: {status_design}")
                if design_force_extraction_available:
                    status_force = "鎴愬姛" if design_force_extraction_successful else "澶辫触"
                    print(f"   {'鉁? if design_force_extraction_successful else '鉂?} 璁捐鍐呭姏鎻愬彇: {status_force}")
                else:
                    print(f"   鈴笍 璁捐鍐呭姏鎻愬彇: 妯″潡涓嶅彲鐢?)
            else:
                print(f"   鈴笍 鏋勪欢璁捐: 妯″潡涓嶅彲鐢?)
                print(f"   鈴笍 璁捐鍐呭姏鎻愬彇: 璺宠繃")
        else:
            print(f"   鈴笍 鏋勪欢璁捐: 璺宠繃")
            print(f"   鈴笍 璁捐鍐呭姏鎻愬彇: 璺宠繃")

        print("=" * 80)

        if not ATTACH_TO_INSTANCE:
            print("鑴氭湰鎵ц瀹屾瘯锛孍TABS 灏嗕繚鎸佹墦寮€鐘舵€佷緵杩涗竴姝ユ搷浣溿€?)

    except SystemExit as e:
        print(f"\n--- 鑴氭湰宸蹭腑姝?---")
        if hasattr(e, 'code') and e.code != 0 and e.code is not None:
            if not (isinstance(e.code, str) and "鍏抽敭閿欒" in e.code):
                print(f"鑴氭湰閫€鍑轰唬鐮? {e.code}")

    except Exception as e:
        print(f"\n--- 鏈鏂欑殑杩愯鏃堕敊璇?---")
        print(f"閿欒绫诲瀷: {type(e).__name__}")
        print(f"閿欒淇℃伅: {e}")
        traceback.print_exc()
        cleanup_etabs_on_error()
        sys.exit(1)

    finally:
        final_elapsed_time = time.time() - script_start_time
        print(f"\n鑴氭湰鎬绘墽琛屾椂闂? {final_elapsed_time:.2f} 绉掋€?)


if __name__ == "__main__":
    main()

