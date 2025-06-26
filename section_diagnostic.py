# -*- coding: utf-8 -*-
"""
API修复脚本 - 解决SetSection参数问题
测试不同的API调用方式
"""

import clr

try:
    clr.AddReference("System")
    import System
except:
    class FakeSystem:
        String, Double = object, object

        class Array:
            @staticmethod
            def CreateInstance(t, s): return []


    System = FakeSystem

from etabs_setup import get_etabs_objects
from config import FRAME_BEAM_SECTION_NAME, FRAME_COLUMN_SECTION_NAME


def test_and_fix_setsection_api():
    """测试并修复SetSection API调用"""
    print("🔧 测试并修复SetSection API")
    print("=" * 50)

    _, sap_model = get_etabs_objects()
    if not sap_model:
        print("❌ 无法连接ETABS")
        return False

    try:
        # 解锁模型
        if sap_model.GetModelIsLocked():
            sap_model.SetModelIsLocked(False)

        # 获取一个测试构件
        NumberNames, FrameNames_tuple = 0, System.Array.CreateInstance(System.String, 0)
        ret, NumberNames, FrameNames_tuple = sap_model.FrameObj.GetNameList(NumberNames, FrameNames_tuple)

        if ret != 0:
            print("❌ 无法获取构件列表")
            return False

        frame_names = list(FrameNames_tuple)
        beam_names = [name for name in frame_names if name.upper().startswith("BEAM")]
        col_names = [name for name in frame_names if name.upper().startswith("COL")]

        if not beam_names:
            print("❌ 没有找到梁构件")
            return False

        test_beam = beam_names[0]
        test_col = col_names[0] if col_names else None

        print(f"📋 测试构件: {test_beam}")
        print(f"📋 目标截面: {FRAME_BEAM_SECTION_NAME}")

        # 测试不同的SetSection调用方式
        success_method = None

        # 方法1: 只传入名称和截面名称
        print("\n🔍 方法1: SetSection(name, section)")
        try:
            ret = sap_model.FrameObj.SetSection(test_beam, FRAME_BEAM_SECTION_NAME)
            print(f"   返回码: {ret}")
            if ret == 0:
                success_method = "method1"
                print("   ✅ 方法1成功!")
        except Exception as e:
            print(f"   ❌ 方法1失败: {e}")

        # 方法2: 传入空字符串作为第三个参数
        if not success_method:
            print("\n🔍 方法2: SetSection(name, section, '')")
            try:
                ret = sap_model.FrameObj.SetSection(test_beam, FRAME_BEAM_SECTION_NAME, "")
                print(f"   返回码: {ret}")
                if ret == 0:
                    success_method = "method2"
                    print("   ✅ 方法2成功!")
            except Exception as e:
                print(f"   ❌ 方法2失败: {e}")

        # 方法3: 使用PropName参数
        if not success_method:
            print("\n🔍 方法3: SetSection(name, section, propname)")
            try:
                ret = sap_model.FrameObj.SetSection(test_beam, FRAME_BEAM_SECTION_NAME, FRAME_BEAM_SECTION_NAME)
                print(f"   返回码: {ret}")
                if ret == 0:
                    success_method = "method3"
                    print("   ✅ 方法3成功!")
            except Exception as e:
                print(f"   ❌ 方法3失败: {e}")

        # 方法4: 尝试获取当前参数然后设置
        if not success_method:
            print("\n🔍 方法4: 先GetSection再SetSection")
            try:
                # 先获取当前设置
                ret_get, current_section, auto_select = sap_model.FrameObj.GetSection(test_beam, "", False)
                print(f"   当前截面: {current_section}, auto: {auto_select}")

                # 然后设置新截面
                ret = sap_model.FrameObj.SetSection(test_beam, FRAME_BEAM_SECTION_NAME, auto_select)
                print(f"   返回码: {ret}")
                if ret == 0:
                    success_method = "method4"
                    print("   ✅ 方法4成功!")
            except Exception as e:
                print(f"   ❌ 方法4失败: {e}")

        # 如果找到成功方法，批量应用
        if success_method:
            print(f"\n🎯 使用成功的方法 {success_method} 批量设置截面...")

            beam_success = 0
            col_success = 0

            # 设置所有梁的截面
            for name in beam_names:
                try:
                    if success_method == "method1":
                        ret = sap_model.FrameObj.SetSection(name, FRAME_BEAM_SECTION_NAME)
                    elif success_method == "method2":
                        ret = sap_model.FrameObj.SetSection(name, FRAME_BEAM_SECTION_NAME, "")
                    elif success_method == "method3":
                        ret = sap_model.FrameObj.SetSection(name, FRAME_BEAM_SECTION_NAME, FRAME_BEAM_SECTION_NAME)
                    elif success_method == "method4":
                        ret_get, current_section, auto_select = sap_model.FrameObj.GetSection(name, "", False)
                        ret = sap_model.FrameObj.SetSection(name, FRAME_BEAM_SECTION_NAME, auto_select)

                    if ret == 0:
                        beam_success += 1
                except:
                    pass

            # 设置所有柱的截面
            for name in col_names:
                try:
                    if success_method == "method1":
                        ret = sap_model.FrameObj.SetSection(name, FRAME_COLUMN_SECTION_NAME)
                    elif success_method == "method2":
                        ret = sap_model.FrameObj.SetSection(name, FRAME_COLUMN_SECTION_NAME, "")
                    elif success_method == "method3":
                        ret = sap_model.FrameObj.SetSection(name, FRAME_COLUMN_SECTION_NAME, FRAME_COLUMN_SECTION_NAME)
                    elif success_method == "method4":
                        ret_get, current_section, auto_select = sap_model.FrameObj.GetSection(name, "", False)
                        ret = sap_model.FrameObj.SetSection(name, FRAME_COLUMN_SECTION_NAME, auto_select)

                    if ret == 0:
                        col_success += 1
                except:
                    pass

            print(f"📊 批量设置结果:")
            print(f"   梁: {beam_success}/{len(beam_names)} 成功")
            print(f"   柱: {col_success}/{len(col_names)} 成功")

            if beam_success > 0 or col_success > 0:
                print("✅ 截面设置成功!")

                # 继续完成设计流程
                return complete_design_workflow(sap_model, beam_names, col_names)
            else:
                print("❌ 批量设置仍然失败")
                return False
        else:
            print("\n❌ 所有SetSection方法都失败了")

            # 尝试检查是否是截面不存在的问题
            print("\n🔍 检查截面是否存在...")
            check_sections_exist(sap_model)
            return False

    except Exception as e:
        print(f"❌ 测试过程出错: {e}")
        return False


def check_sections_exist(sap_model):
    """检查截面是否存在"""
    try:
        num_names = 0
        section_names = System.Array.CreateInstance(System.String, 0)
        ret, num_names, section_names = sap_model.PropFrame.GetNameList(num_names, section_names)

        if ret == 0:
            sections = list(section_names)
            print(f"   模型中的所有截面 ({len(sections)} 个):")
            for section in sections:
                print(f"     - {section}")

            if FRAME_BEAM_SECTION_NAME not in sections:
                print(f"   ❌ 梁截面 {FRAME_BEAM_SECTION_NAME} 不存在!")
            if FRAME_COLUMN_SECTION_NAME not in sections:
                print(f"   ❌ 柱截面 {FRAME_COLUMN_SECTION_NAME} 不存在!")
    except Exception as e:
        print(f"   检查截面失败: {e}")


def complete_design_workflow(sap_model, beam_names, col_names):
    """完成设计工作流"""
    print("\n🚀 完成设计工作流...")

    try:
        # 创建分组
        for group_name in ["ALL_BEAMS", "ALL_COLUMNS"]:
            try:
                sap_model.GroupDef.Delete(group_name)
            except:
                pass
            sap_model.GroupDef.SetGroup(group_name)

        # 分组构件
        beam_grouped = 0
        for name in beam_names:
            try:
                ret = sap_model.FrameObj.SetGroupAssign(name, "ALL_BEAMS", True)
                if ret == 0:
                    beam_grouped += 1
            except:
                pass

        col_grouped = 0
        for name in col_names:
            try:
                ret = sap_model.FrameObj.SetGroupAssign(name, "ALL_COLUMNS", True)
                if ret == 0:
                    col_grouped += 1
            except:
                pass

        print(f"   分组: 梁 {beam_grouped}/{len(beam_names)}, 柱 {col_grouped}/{len(col_names)}")

        # 设置混凝土设计
        all_frames = beam_names + col_names
        design_set = 0
        for name in all_frames:
            try:
                ret = sap_model.FrameObj.SetDesignProcedure(name, 2)
                if ret == 0:
                    design_set += 1
            except:
                pass

        print(f"   设计程序: {design_set}/{len(all_frames)}")

        # 保存并分析
        sap_model.File.Save()
        sap_model.SetModelIsLocked(True)
        ret = sap_model.Analyze.RunAnalysis()
        print(f"   分析: {'✅' if ret == 0 else '❌'}")

        # 运行设计
        try:
            sap_model.DesignConcrete.SetCode("Chinese 2010")
        except:
            pass

        ret = sap_model.DesignConcrete.StartDesign()
        print(f"   设计: {'✅' if ret == 0 else '❌'} (返回码: {ret})")

        if ret == 0:
            print("🎉 设计成功完成!")

            # 测试提取结果
            test_extract_results(sap_model, beam_names[:1])
            return True
        else:
            print("❌ 设计失败")
            return False

    except Exception as e:
        print(f"   工作流失败: {e}")
        return False


def test_extract_results(sap_model, test_beams):
    """测试结果提取"""
    if not test_beams:
        return

    print("📊 测试结果提取...")
    try:
        dc = sap_model.DesignConcrete
        test_beam = test_beams[0]

        num_items = 0
        obj_names = System.Array.CreateInstance(System.String, 0)
        elmn_names = System.Array.CreateInstance(System.String, 0)
        load_cases = System.Array.CreateInstance(System.String, 0)
        locations = System.Array.CreateInstance(System.Double, 0)
        top_areas = System.Array.CreateInstance(System.Double, 0)
        bot_areas = System.Array.CreateInstance(System.Double, 0)

        res = dc.GetSummaryResultsBeam(
            test_beam, num_items, obj_names, elmn_names, load_cases,
            locations, top_areas, bot_areas,
            System.Enum.ToObject(dc.GetType().Module.GetType("ETABSv1.eItemType"), 0)
        )

        print(f"   测试梁: {test_beam}")
        print(f"   API返回: 码={res[0]}, 结果数={res[1]}")

        if res[0] == 0 and res[1] > 0:
            top_areas_list = list(res[5]) if len(res) > 5 else []
            bot_areas_list = list(res[6]) if len(res) > 6 else []

            top_max = max([a * 1e6 for a in top_areas_list if a is not None and a > 0], default=0)
            bot_max = max([a * 1e6 for a in bot_areas_list if a is not None and a > 0], default=0)

            print(f"   配筋结果: 上部 {top_max:.2f} mm², 下部 {bot_max:.2f} mm²")

            if top_max > 0 or bot_max > 0:
                print("   ✅ 结果提取成功!")
            else:
                print("   ⚠️ 配筋为0，可能设计未运行")
        else:
            print("   ❌ 无设计结果")

    except Exception as e:
        print(f"   提取失败: {e}")


if __name__ == "__main__":
    success = test_and_fix_setsection_api()
    if success:
        print("\n🎉 API修复成功!")
    else:
        print("\n❌ API修复失败")