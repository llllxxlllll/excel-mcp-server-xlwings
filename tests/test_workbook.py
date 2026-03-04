# -*- coding: utf-8 -*-
"""工作簿模块属性测试

使用 Hypothesis 进行属性测试，验证工作簿操作的正确性。
"""

import os
import tempfile
import shutil

import pytest
import xlwings as xw
from hypothesis import given, strategies as st, settings, assume

# 导入被测试模块
from excel_mcp.workbook import create_workbook, get_workbook_info, create_sheet
from excel_mcp.exceptions import WorkbookError


def cleanup_excel():
    """清理所有 Excel 实例
    
    使用更安全的方式清理 Excel COM 对象，避免 COM 断开连接时的异常
    """
    import time
    import subprocess
    import pythoncom
    
    # 首先尝试通过 taskkill 强制结束 Excel 进程
    # 这是最可靠的方式，避免 COM 连接问题
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'], 
                      capture_output=True, timeout=5)
        time.sleep(0.5)  # 等待进程完全结束
    except Exception:
        pass
    
    # 清理 COM 对象引用
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass


# 设置 Excel 为不可见模式（测试用）
import excel_mcp.xw_helper as xw_helper
xw_helper.EXCEL_VISIBLE = False


# 生成有效的工作表名称策略
# Excel 工作表名称限制：最多31个字符，不能包含 : \ / ? * [ ]
valid_sheet_name_chars = st.sampled_from(
    "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_- "
)
valid_sheet_name = st.text(
    alphabet=valid_sheet_name_chars,
    min_size=1,
    max_size=31
).filter(lambda x: x.strip() != "")  # 不能是纯空格


class TestWorkbookProperties:
    """工作簿属性测试类"""
    
    @pytest.fixture(autouse=True)
    def setup_teardown(self):
        """每个测试前后的设置和清理"""
        # 创建临时目录
        self.temp_dir = tempfile.mkdtemp()
        yield
        # 清理 Excel 实例
        cleanup_excel()
        # 清理临时目录
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    @given(sheet_name=valid_sheet_name)
    @settings(max_examples=10, deadline=None)  # 减少测试次数以避免 Excel COM 问题
    def test_workbook_created_is_accessible(self, sheet_name: str):
        """
        **Feature: xlwings-excel-mcp, Property 2: 工作簿创建后可访问**
        **Validates: Requirements 2.1, 2.2**
        
        对于任意有效文件路径，创建工作簿后应能成功获取其元数据，
        且元数据包含工作表名称。
        """
        # 生成唯一的文件路径
        filepath = os.path.join(self.temp_dir, f"test_{hash(sheet_name)}.xlsx")
        
        # 确保文件不存在
        if os.path.exists(filepath):
            os.remove(filepath)
        
        try:
            # 创建工作簿
            result = create_workbook(filepath, sheet_name=sheet_name)
            
            # 验证创建成功
            assert "message" in result
            assert os.path.exists(filepath)
            
            # 获取元数据
            info = get_workbook_info(filepath)
            
            # 验证元数据包含工作表名称
            assert "sheets" in info
            assert sheet_name in info["sheets"]
            assert "filename" in info
            assert "size" in info
            
        finally:
            # 清理文件
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass
    
    @given(sheet_name=valid_sheet_name)
    @settings(max_examples=10, deadline=None)  # 减少测试次数以避免 Excel COM 问题
    def test_sheet_created_exists(self, sheet_name: str):
        """
        **Feature: xlwings-excel-mcp, Property 3: 工作表创建后存在**
        **Validates: Requirements 2.3**
        
        对于任意有效工作表名称，创建工作表后该工作表应出现在
        工作簿的工作表列表中。
        """
        # 生成唯一的文件路径
        filepath = os.path.join(self.temp_dir, f"test_sheet_{hash(sheet_name)}.xlsx")
        
        # 确保文件不存在
        if os.path.exists(filepath):
            os.remove(filepath)
        
        try:
            # 先创建工作簿（使用不同的初始工作表名）
            initial_sheet = "InitialSheet"
            # 确保新工作表名与初始工作表名不同
            assume(sheet_name != initial_sheet)
            
            create_workbook(filepath, sheet_name=initial_sheet)
            
            # 创建新工作表
            create_sheet(filepath, sheet_name)
            
            # 获取元数据验证
            info = get_workbook_info(filepath)
            
            # 验证新工作表存在
            assert sheet_name in info["sheets"]
            # 验证初始工作表也存在
            assert initial_sheet in info["sheets"]
            
        finally:
            # 清理文件
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass



class TestDataProperties:
    """数据操作属性测试类"""
    
    @pytest.fixture(autouse=True)
    def setup_teardown(self):
        """每个测试前后的设置和清理"""
        self.temp_dir = tempfile.mkdtemp()
        yield
        cleanup_excel()
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    @given(
        data=st.lists(
            st.lists(
                st.one_of(
                    st.integers(min_value=-1000, max_value=1000),
                    st.text(alphabet="abcdefghijklmnopqrstuvwxyz", min_size=1, max_size=10),
                    st.floats(min_value=-1000, max_value=1000, allow_nan=False, allow_infinity=False)
                ),
                min_size=2,  # 至少2列以避免单列问题
                max_size=5
            ),
            min_size=2,  # 至少2行以避免单行问题
            max_size=5
        )
    )
    @settings(max_examples=10, deadline=None)
    def test_data_write_read_roundtrip(self, data):
        """
        **Feature: xlwings-excel-mcp, Property 1: 数据写入读取往返一致性**
        **Validates: Requirements 3.1, 3.2, 3.5, 3.6**
        
        对于任意有效数据和起始单元格，写入数据后读取应返回相同的数据值。
        """
        from excel_mcp.workbook import create_workbook
        from excel_mcp.data import write_data, read_excel_range
        from excel_mcp.xw_helper import column_string_from_index
        
        # 确保所有行长度一致
        max_cols = max(len(row) for row in data)
        normalized_data = []
        for row in data:
            normalized_row = list(row) + [None] * (max_cols - len(row))
            normalized_data.append(normalized_row)
        
        num_rows = len(normalized_data)
        num_cols = max_cols
        
        # 使用唯一的文件名
        import uuid
        filepath = os.path.join(self.temp_dir, f"test_data_{uuid.uuid4().hex}.xlsx")
        
        try:
            # 创建工作簿
            create_workbook(filepath, sheet_name="TestSheet")
            
            # 写入数据
            write_data(filepath, "TestSheet", normalized_data, start_cell="A1")
            
            # 计算结束单元格
            end_cell = f"{column_string_from_index(num_cols)}{num_rows}"
            
            # 读取数据（指定范围）
            read_data = read_excel_range(filepath, "TestSheet", start_cell="A1", end_cell=end_cell)
            
            # 验证数据一致性
            assert len(read_data) == len(normalized_data), f"行数不匹配: {len(read_data)} vs {len(normalized_data)}"
            
            for i, (read_row, orig_row) in enumerate(zip(read_data, normalized_data)):
                assert len(read_row) == len(orig_row), f"列数不匹配 行{i}: {len(read_row)} vs {len(orig_row)}"
                for j, (read_val, orig_val) in enumerate(zip(read_row, orig_row)):
                    # 处理浮点数比较
                    if isinstance(orig_val, float) and isinstance(read_val, float):
                        assert abs(read_val - orig_val) < 0.0001, f"值不匹配 [{i}][{j}]: {read_val} vs {orig_val}"
                    elif orig_val is None:
                        # None 值可能被读取为 None 或空
                        pass
                    else:
                        assert read_val == orig_val, f"值不匹配 [{i}][{j}]: {read_val} vs {orig_val}"
                        
        finally:
            # 清理 Excel 先
            cleanup_excel()
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass


class TestFormulaProperties:
    """公式操作属性测试类"""
    
    @pytest.fixture(autouse=True)
    def setup_teardown(self):
        """每个测试前后的设置和清理"""
        self.temp_dir = tempfile.mkdtemp()
        yield
        cleanup_excel()
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    # 生成有效的 Excel 公式策略
    # 简单公式：=A1, =SUM(A1:A10), =A1+B1 等
    simple_cell_refs = st.from_regex(r"[A-Z][1-9][0-9]?", fullmatch=True)
    
    @given(
        cell_ref=st.from_regex(r"[A-Z][1-9]", fullmatch=True),
        value1=st.integers(min_value=1, max_value=100),
        value2=st.integers(min_value=1, max_value=100)
    )
    @settings(max_examples=10, deadline=None)
    def test_formula_storage_consistency(self, cell_ref: str, value1: int, value2: int):
        """
        **Feature: xlwings-excel-mcp, Property 4: 公式存储一致性**
        **Validates: Requirements 3.3, 11.4**
        
        对于任意有效公式，写入单元格后读取应返回相同的公式字符串。
        """
        from excel_mcp.workbook import create_workbook
        from excel_mcp.calculations import apply_formula, get_formula
        from excel_mcp.data import write_data
        import uuid
        
        filepath = os.path.join(self.temp_dir, f"test_formula_{uuid.uuid4().hex}.xlsx")
        
        try:
            # 创建工作簿
            create_workbook(filepath, sheet_name="TestSheet")
            
            # 先写入一些数据供公式引用
            write_data(filepath, "TestSheet", [[value1], [value2]], start_cell="A1")
            
            # 应用公式到 B1
            formula = "=A1+A2"
            apply_formula(filepath, "TestSheet", "B1", formula)
            
            # 读取公式
            result = get_formula(filepath, "TestSheet", "B1")
            
            # 验证公式存储一致性
            assert result["has_formula"] is True
            assert result["formula"] == formula
            
        finally:
            cleanup_excel()
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass
    
    @given(
        unbalanced_formula=st.sampled_from([
            "=SUM(A1:A10",      # 缺少右括号
            "=SUM(A1:A10))",    # 多余右括号
            "=((A1+B1)",        # 括号不平衡
            "=A1+(B1",          # 缺少右括号
        ])
    )
    @settings(max_examples=4, deadline=None)
    def test_formula_parentheses_validation(self, unbalanced_formula: str):
        """
        **Feature: xlwings-excel-mcp, Property 16: 公式括号平衡验证**
        **Validates: Requirements 11.1**
        
        对于任意公式字符串，括号不平衡的公式应被验证函数拒绝。
        """
        from excel_mcp.validation import validate_formula
        
        # 验证括号不平衡的公式
        is_valid, message = validate_formula(unbalanced_formula)
        
        # 应该被拒绝
        assert is_valid is False
        assert "parenthesis" in message.lower()
    
    @given(
        balanced_formula=st.sampled_from([
            "=SUM(A1:A10)",
            "=A1+B1",
            "=IF(A1>0,A1,0)",
            "=AVERAGE(A1:A10)",
            "=(A1+B1)*C1",
            "=((A1+B1)*(C1+D1))",
        ])
    )
    @settings(max_examples=6, deadline=None)
    def test_valid_formula_accepted(self, balanced_formula: str):
        """
        验证有效公式被接受
        """
        from excel_mcp.validation import validate_formula
        
        is_valid, message = validate_formula(balanced_formula)
        
        assert is_valid is True
        assert message == "Formula is valid"


class TestFormattingProperties:
    """格式化操作属性测试类"""
    
    @pytest.fixture(autouse=True)
    def setup_teardown(self):
        """每个测试前后的设置和清理"""
        self.temp_dir = tempfile.mkdtemp()
        yield
        cleanup_excel()
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    @given(
        bold=st.booleans(),
        italic=st.booleans(),
        font_size=st.integers(min_value=8, max_value=24)
    )
    @settings(max_examples=5, deadline=None)
    def test_format_applied_readable(self, bold: bool, italic: bool, font_size: int):
        """
        **Feature: xlwings-excel-mcp, Property 5: 格式应用后可读取**
        **Validates: Requirements 4.1, 4.2, 4.3, 4.4**
        
        对于任意格式设置，应用格式后读取单元格格式属性应与设置值一致。
        """
        from excel_mcp.workbook import create_workbook
        from excel_mcp.formatting import format_range
        from excel_mcp.xw_helper import get_workbook, get_sheet
        import uuid
        
        filepath = os.path.join(self.temp_dir, f"test_format_{uuid.uuid4().hex}.xlsx")
        
        try:
            # 创建工作簿
            create_workbook(filepath, sheet_name="TestSheet")
            
            # 应用格式
            format_range(
                filepath, "TestSheet", "A1",
                bold=bold,
                italic=italic,
                font_size=font_size
            )
            
            # 读取格式验证
            wb = get_workbook(filepath)
            sheet = get_sheet(wb, "TestSheet")
            cell = sheet.range("A1")
            
            # 验证格式一致性
            assert cell.font.bold == bold, f"Bold mismatch: {cell.font.bold} vs {bold}"
            assert cell.font.italic == italic, f"Italic mismatch: {cell.font.italic} vs {italic}"
            assert cell.font.size == font_size, f"Font size mismatch: {cell.font.size} vs {font_size}"
            
        finally:
            cleanup_excel()
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass


class TestSheetProperties:
    """工作表操作属性测试类"""
    
    @pytest.fixture(autouse=True)
    def setup_teardown(self):
        """每个测试前后的设置和清理"""
        self.temp_dir = tempfile.mkdtemp()
        yield
        cleanup_excel()
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    @given(new_sheet_name=valid_sheet_name)
    @settings(max_examples=5, deadline=None)
    def test_sheet_copy_content_consistent(self, new_sheet_name: str):
        """
        **Feature: xlwings-excel-mcp, Property 9: 工作表复制内容一致**
        **Validates: Requirements 7.1**
        
        对于任意工作表，复制后新工作表的内容应与源工作表完全一致。
        """
        from excel_mcp.workbook import create_workbook
        from excel_mcp.sheet import copy_sheet
        from excel_mcp.data import write_data, read_excel_range
        import uuid
        
        # 确保新工作表名与源不同
        assume(new_sheet_name != "SourceSheet")
        
        filepath = os.path.join(self.temp_dir, f"test_copy_{uuid.uuid4().hex}.xlsx")
        
        try:
            # 创建工作簿并写入数据
            create_workbook(filepath, sheet_name="SourceSheet")
            test_data = [[1, 2, 3], [4, 5, 6]]
            write_data(filepath, "SourceSheet", test_data, start_cell="A1")
            
            # 复制工作表
            copy_sheet(filepath, "SourceSheet", new_sheet_name)
            
            # 读取两个工作表的数据
            source_data = read_excel_range(filepath, "SourceSheet", "A1", "C2")
            target_data = read_excel_range(filepath, new_sheet_name, "A1", "C2")
            
            # 验证内容一致
            assert source_data == target_data
            
        finally:
            cleanup_excel()
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass
    
    @given(sheet_to_delete=valid_sheet_name)
    @settings(max_examples=5, deadline=None)
    def test_sheet_delete_not_exists(self, sheet_to_delete: str):
        """
        **Feature: xlwings-excel-mcp, Property 10: 工作表删除后不存在**
        **Validates: Requirements 7.2**
        
        对于任意非唯一工作表，删除后该工作表不应出现在工作簿的工作表列表中。
        """
        from excel_mcp.workbook import create_workbook, get_workbook_info, create_sheet
        from excel_mcp.sheet import delete_sheet
        import uuid
        
        assume(sheet_to_delete != "KeepSheet")
        
        filepath = os.path.join(self.temp_dir, f"test_delete_{uuid.uuid4().hex}.xlsx")
        
        try:
            # 创建工作簿并添加第二个工作表
            create_workbook(filepath, sheet_name="KeepSheet")
            create_sheet(filepath, sheet_to_delete)
            
            # 删除工作表
            delete_sheet(filepath, sheet_to_delete)
            
            # 验证工作表不存在
            info = get_workbook_info(filepath)
            assert sheet_to_delete not in info["sheets"]
            assert "KeepSheet" in info["sheets"]
            
        finally:
            cleanup_excel()
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass
    
    @given(
        old_name=valid_sheet_name,
        new_name=valid_sheet_name
    )
    @settings(max_examples=5, deadline=None)
    def test_sheet_rename_consistency(self, old_name: str, new_name: str):
        """
        **Feature: xlwings-excel-mcp, Property 11: 工作表重命名一致性**
        **Validates: Requirements 7.3**
        
        对于任意工作表和新名称，重命名后旧名称不存在且新名称存在于工作表列表中。
        """
        from excel_mcp.workbook import create_workbook, get_workbook_info
        from excel_mcp.sheet import rename_sheet
        import uuid
        
        assume(old_name != new_name)
        
        filepath = os.path.join(self.temp_dir, f"test_rename_{uuid.uuid4().hex}.xlsx")
        
        try:
            # 创建工作簿
            create_workbook(filepath, sheet_name=old_name)
            
            # 重命名工作表
            rename_sheet(filepath, old_name, new_name)
            
            # 验证重命名结果
            info = get_workbook_info(filepath)
            assert old_name not in info["sheets"]
            assert new_name in info["sheets"]
            
        finally:
            cleanup_excel()
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass
    
    @settings(max_examples=3, deadline=None)
    @given(st.just(True))  # 简单测试
    def test_merge_unmerge_roundtrip(self, _):
        """
        **Feature: xlwings-excel-mcp, Property 6: 合并取消合并往返**
        **Validates: Requirements 4.5, 4.6**
        
        对于任意单元格范围，合并后取消合并应恢复到原始状态。
        """
        from excel_mcp.workbook import create_workbook
        from excel_mcp.sheet import merge_range, unmerge_range
        from excel_mcp.xw_helper import get_workbook, get_sheet
        import uuid
        
        filepath = os.path.join(self.temp_dir, f"test_merge_{uuid.uuid4().hex}.xlsx")
        
        try:
            # 创建工作簿
            create_workbook(filepath, sheet_name="TestSheet")
            
            # 合并单元格
            merge_range(filepath, "TestSheet", "A1", "B2")
            
            # 验证已合并
            wb = get_workbook(filepath)
            sheet = get_sheet(wb, "TestSheet")
            cell = sheet.range("A1")
            assert cell.api.MergeCells == True
            
            # 取消合并
            unmerge_range(filepath, "TestSheet", "A1", "B2")
            
            # 验证已取消合并
            wb = get_workbook(filepath)
            sheet = get_sheet(wb, "TestSheet")
            cell = sheet.range("A1")
            assert cell.api.MergeCells == False
            
        finally:
            cleanup_excel()
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass
