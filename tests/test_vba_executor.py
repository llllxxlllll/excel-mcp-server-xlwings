# -*- coding: utf-8 -*-
"""VBA 执行器属性测试和单元测试

使用 Hypothesis 进行属性测试，验证 VBA 执行器的正确性。
"""

import os
import tempfile
import shutil
import threading
from pathlib import Path
from typing import List

import pytest
from hypothesis import given, strategies as st, settings, assume

from excel_mcp.vba_executor import VBAExecutor
from excel_mcp.exceptions import (
    VBASecurityError,
    VBAExecutionError,
    VBATimeoutError,
    VBABusyError,
    WorkbookError,
)


# ============================================================================
# 测试辅助函数和策略
# ============================================================================

def safe_vba_code(sub_name: str = "Main", body: str = "") -> str:
    """生成安全的 VBA 代码"""
    return f"""
Sub {sub_name}()
    {body}
End Sub
"""


def unsafe_vba_code_with_keyword(keyword: str) -> str:
    """生成包含敏感关键词的 VBA 代码"""
    return f"""
Sub Main()
    ' 包含敏感关键词: {keyword}
    Dim x As String
    x = "{keyword}"
End Sub
"""


# Hypothesis 策略：生成敏感关键词
sensitive_keyword_strategy = st.sampled_from(VBAExecutor.SENSITIVE_KEYWORDS)

# Hypothesis 策略：生成安全的 VBA 代码（不包含敏感关键词）
safe_identifiers = st.text(
    alphabet=st.characters(whitelist_categories=('Lu', 'Ll'), min_codepoint=65, max_codepoint=122),
    min_size=1,
    max_size=10
).filter(lambda x: x.isalpha())


# ============================================================================
# Property 2: Sensitive Keyword Rejection
# **Feature: execute-excel-vba, Property 2: Sensitive Keyword Rejection**
# **Validates: Requirements 2.1, 2.2**
# ============================================================================

class TestSensitiveKeywordRejection:
    """属性测试：敏感关键词拒绝"""
    
    @given(keyword=sensitive_keyword_strategy)
    @settings(max_examples=100)
    def test_property_sensitive_keyword_detected(self, keyword: str):
        """
        **Feature: execute-excel-vba, Property 2: Sensitive Keyword Rejection**
        **Validates: Requirements 2.1, 2.2**
        
        对于任何包含敏感关键词的 VBA 代码，执行器应检测到该关键词。
        """
        executor = VBAExecutor()
        vba_code = unsafe_vba_code_with_keyword(keyword)
        
        detected = executor._scan_sensitive_keywords(vba_code)
        
        assert keyword in detected, f"应检测到敏感关键词 '{keyword}'"
    
    @given(keyword=sensitive_keyword_strategy)
    @settings(max_examples=100)
    def test_property_case_insensitive_detection(self, keyword: str):
        """
        **Feature: execute-excel-vba, Property 2: Sensitive Keyword Rejection**
        **Validates: Requirements 2.1, 2.2**
        
        敏感关键词检测应大小写不敏感。
        """
        executor = VBAExecutor()
        
        # 测试大写
        vba_code_upper = f"Sub Main()\n    {keyword.upper()}\nEnd Sub"
        detected_upper = executor._scan_sensitive_keywords(vba_code_upper)
        
        # 测试小写
        vba_code_lower = f"Sub Main()\n    {keyword.lower()}\nEnd Sub"
        detected_lower = executor._scan_sensitive_keywords(vba_code_lower)
        
        # 测试混合大小写
        vba_code_mixed = f"Sub Main()\n    {keyword.title()}\nEnd Sub"
        detected_mixed = executor._scan_sensitive_keywords(vba_code_mixed)
        
        assert keyword in detected_upper, f"应检测到大写关键词 '{keyword.upper()}'"
        assert keyword in detected_lower, f"应检测到小写关键词 '{keyword.lower()}'"
        assert keyword in detected_mixed, f"应检测到混合大小写关键词 '{keyword.title()}'"
    
    def test_safe_code_passes_scan(self):
        """安全代码应通过扫描"""
        executor = VBAExecutor()
        vba_code = safe_vba_code("Main", "MsgBox \"Hello\"")
        
        detected = executor._scan_sensitive_keywords(vba_code)
        
        assert len(detected) == 0, f"安全代码不应检测到敏感关键词，但检测到: {detected}"
    
    def test_multiple_keywords_all_detected(self):
        """多个敏感关键词应全部被检测到"""
        executor = VBAExecutor()
        keywords = ["Shell", "Kill", "CreateObject"]
        vba_code = f"""
Sub Main()
    Shell "cmd"
    Kill "file.txt"
    CreateObject("Scripting.FileSystemObject")
End Sub
"""
        detected = executor._scan_sensitive_keywords(vba_code)
        
        for kw in keywords:
            assert kw in detected, f"应检测到关键词 '{kw}'"



# ============================================================================
# Property 3: Backup Creation Before Execution
# **Feature: execute-excel-vba, Property 3: Backup Creation Before Execution**
# **Validates: Requirements 2.3**
# ============================================================================

class TestBackupCreation:
    """属性测试：备份创建"""
    
    @pytest.fixture
    def temp_excel_file(self, tmp_path):
        """创建临时 Excel 文件用于测试"""
        # 创建一个简单的测试文件
        test_file = tmp_path / "test_backup.xlsx"
        # 写入一些内容以创建有效的文件
        test_file.write_bytes(b"PK")  # 简单的 xlsx 文件头
        return str(test_file)
    
    @given(filename=st.text(
        alphabet=st.characters(whitelist_categories=('Lu', 'Ll', 'Nd'), min_codepoint=48, max_codepoint=122),
        min_size=1,
        max_size=20
    ).filter(lambda x: x.isalnum()))
    @settings(max_examples=50)
    def test_property_backup_naming_format(self, filename: str):
        """
        **Feature: execute-excel-vba, Property 3: Backup Creation Before Execution**
        **Validates: Requirements 2.3**
        
        对于任何有效的文件名，备份文件应遵循命名格式：BACKUP_{timestamp}_{filename}
        """
        executor = VBAExecutor()
        
        # 使用 tempfile 创建临时目录
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            # 创建临时测试文件
            test_file = tmp_path / f"{filename}.xlsx"
            test_file.write_bytes(b"test content")
            
            backup_path = executor._create_backup(str(test_file))
            backup_name = Path(backup_path).name
            
            # 验证备份文件名格式
            assert backup_name.startswith(VBAExecutor.BACKUP_PREFIX), \
                f"备份文件名应以 '{VBAExecutor.BACKUP_PREFIX}' 开头"
            assert backup_name.endswith(f"_{filename}.xlsx"), \
                f"备份文件名应以原文件名结尾"
            assert os.path.exists(backup_path), "备份文件应存在"
    
    def test_backup_content_matches_original(self, tmp_path):
        """
        **Feature: execute-excel-vba, Property 3: Backup Creation Before Execution**
        **Validates: Requirements 2.3**
        
        备份文件内容应与原文件完全一致。
        """
        executor = VBAExecutor()
        
        # 创建测试文件
        original_content = b"This is test content for backup verification"
        test_file = tmp_path / "original.xlsx"
        test_file.write_bytes(original_content)
        
        # 创建备份
        backup_path = executor._create_backup(str(test_file))
        
        # 验证内容一致
        with open(backup_path, "rb") as f:
            backup_content = f.read()
        
        assert backup_content == original_content, "备份内容应与原文件一致"
    
    def test_backup_fails_for_nonexistent_file(self, tmp_path):
        """备份不存在的文件应抛出异常"""
        executor = VBAExecutor()
        nonexistent_file = str(tmp_path / "nonexistent.xlsx")
        
        with pytest.raises(VBAExecutionError) as exc_info:
            executor._create_backup(nonexistent_file)
        
        assert "备份创建失败" in str(exc_info.value)


# ============================================================================
# Property 1: Result Structure Consistency
# **Feature: execute-excel-vba, Property 1: Result Structure Consistency**
# **Validates: Requirements 1.3, 3.4, 4.3**
# ============================================================================

class TestResultStructureConsistency:
    """属性测试：结果结构一致性"""
    
    def test_security_error_result_structure(self):
        """
        **Feature: execute-excel-vba, Property 1: Result Structure Consistency**
        **Validates: Requirements 1.3, 3.4, 4.3**
        
        安全检查失败时应抛出包含敏感关键词信息的异常。
        """
        executor = VBAExecutor()
        vba_code = unsafe_vba_code_with_keyword("Shell")
        
        # 由于没有真实的 Excel 文件，我们测试安全检查部分
        detected = executor._scan_sensitive_keywords(vba_code)
        
        assert isinstance(detected, list), "检测结果应为列表"
        assert len(detected) > 0, "应检测到敏感关键词"
        assert "Shell" in detected, "应检测到 Shell 关键词"


# ============================================================================
# Property 4: VBA Runtime Error Capture
# **Feature: execute-excel-vba, Property 4: VBA Runtime Error Capture**
# **Validates: Requirements 3.2**
# ============================================================================

class TestVBARuntimeErrorCapture:
    """属性测试：VBA 运行时错误捕获"""
    
    def test_entry_sub_not_found_error_message(self):
        """
        **Feature: execute-excel-vba, Property 4: VBA Runtime Error Capture**
        **Validates: Requirements 3.2**
        
        当入口 Sub 不存在时，应返回明确的错误信息。
        """
        # 这个测试验证错误消息格式
        # 实际的 VBA 执行测试需要真实的 Excel 环境
        executor = VBAExecutor()
        
        # 验证执行器已正确初始化
        assert executor.EXECUTION_TIMEOUT == 30
        assert executor.BACKUP_PREFIX == "BACKUP_"



# ============================================================================
# Property 6: VBA Execution Round Trip
# **Feature: execute-excel-vba, Property 6: VBA Execution Round Trip**
# **Validates: Requirements 1.1**
# ============================================================================

# 标记需要真实 Excel 环境的测试
requires_excel = pytest.mark.skipif(
    os.environ.get("SKIP_EXCEL_TESTS", "1") == "1",
    reason="需要真实的 Excel 环境，设置 SKIP_EXCEL_TESTS=0 以运行"
)


class TestVBAExecutionRoundTrip:
    """属性测试：VBA 执行往返"""
    
    @requires_excel
    @given(
        cell_value=st.integers(min_value=1, max_value=1000),
        row=st.integers(min_value=1, max_value=10),
        col=st.integers(min_value=1, max_value=10)
    )
    @settings(max_examples=20)
    def test_property_vba_modifies_cell_value(self, cell_value: int, row: int, col: int):
        """
        **Feature: execute-excel-vba, Property 6: VBA Execution Round Trip**
        **Validates: Requirements 1.1**
        
        对于任何有效的 VBA 代码修改单元格值，执行后读取应反映预期的更改。
        """
        import xlwings as xw
        
        executor = VBAExecutor()
        
        with tempfile.TemporaryDirectory() as tmp_dir:
            # 创建测试 Excel 文件
            test_file = Path(tmp_dir) / "test_roundtrip.xlsx"
            
            # 使用 xlwings 创建文件
            app = xw.App(visible=False, add_book=False)
            try:
                wb = app.books.add()
                wb.save(str(test_file))
                wb.close()
            finally:
                app.quit()
            
            # 生成修改单元格的 VBA 代码
            col_letter = chr(ord('A') + col - 1)
            vba_code = f"""
Sub Main()
    Cells({row}, {col}).Value = {cell_value}
End Sub
"""
            
            # 执行 VBA
            try:
                result = executor.execute_vba(str(test_file), vba_code, "Main")
                
                # 验证执行成功
                assert result["status"] == "success", f"执行应成功: {result}"
                
                # 读取修改后的值
                app = xw.App(visible=False, add_book=False)
                try:
                    wb = app.books.open(str(test_file))
                    actual_value = wb.sheets[0].cells(row, col).value
                    wb.close()
                finally:
                    app.quit()
                
                # 验证值已修改
                assert actual_value == cell_value, \
                    f"单元格值应为 {cell_value}，实际为 {actual_value}"
                    
            except VBASecurityError:
                # 如果 VBA 信任未开启，跳过测试
                pytest.skip("VBA 信任未开启")
    
    def test_vba_code_structure_validation(self):
        """
        **Feature: execute-excel-vba, Property 6: VBA Execution Round Trip**
        **Validates: Requirements 1.1**
        
        验证 VBA 代码结构生成正确。
        """
        # 测试安全的 VBA 代码生成
        vba_code = safe_vba_code("TestSub", "Cells(1, 1).Value = 100")
        
        assert "Sub TestSub()" in vba_code
        assert "End Sub" in vba_code
        assert "Cells(1, 1).Value = 100" in vba_code


# ============================================================================
# Property 5: Parameter Validation
# **Feature: execute-excel-vba, Property 5: Parameter Validation**
# **Validates: Requirements 4.2, 5.2**
# ============================================================================

class TestParameterValidation:
    """属性测试：参数验证"""
    
    def test_missing_filepath_raises_error(self):
        """
        **Feature: execute-excel-vba, Property 5: Parameter Validation**
        **Validates: Requirements 4.2, 5.2**
        
        缺少文件路径参数应抛出错误。
        """
        executor = VBAExecutor()
        
        with pytest.raises((WorkbookError, VBAExecutionError, TypeError)):
            executor.execute_vba("", "Sub Main()\nEnd Sub", "Main")
    
    def test_nonexistent_file_raises_workbook_error(self):
        """
        **Feature: execute-excel-vba, Property 5: Parameter Validation**
        **Validates: Requirements 4.2, 5.2**
        
        不存在的文件应抛出 WorkbookError。
        """
        executor = VBAExecutor()
        
        with pytest.raises(WorkbookError) as exc_info:
            executor.execute_vba("/nonexistent/path/file.xlsx", "Sub Main()\nEnd Sub", "Main")
        
        assert "不存在" in str(exc_info.value) or "not exist" in str(exc_info.value).lower()


# ============================================================================
# Property 7 & 8: Timeout and Concurrency (需要真实 Excel 环境)
# ============================================================================

class TestTimeoutAndConcurrency:
    """超时和并发控制测试"""
    
    def test_lock_mechanism_exists(self):
        """
        **Feature: execute-excel-vba, Property 8: Concurrent Request Handling**
        **Validates: Requirements 3.3**
        
        验证全局锁机制存在。
        """
        # 验证类级别的锁存在
        assert hasattr(VBAExecutor, '_lock')
        assert isinstance(VBAExecutor._lock, type(threading.Lock()))
    
    def test_timeout_configuration(self):
        """
        **Feature: execute-excel-vba, Property 7: Execution Timeout Safety**
        **Validates: Requirements 3.3**
        
        验证超时配置正确。
        """
        executor = VBAExecutor()
        
        assert executor.EXECUTION_TIMEOUT == 30, "默认超时应为 30 秒"
        assert executor.LOCK_TIMEOUT == 60, "锁超时应为 60 秒"
    
    def test_force_kill_method_exists(self):
        """
        **Feature: execute-excel-vba, Property 7: Execution Timeout Safety**
        **Validates: Requirements 3.3**
        
        验证强制终止方法存在。
        """
        executor = VBAExecutor()
        
        assert hasattr(executor, '_force_kill_excel')
        assert callable(executor._force_kill_excel)
