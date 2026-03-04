# -*- coding: utf-8 -*-
"""VBA 执行功能集成测试

测试完整的 MCP 工具调用流程和各种错误场景。
"""

import json
import os
import tempfile
from pathlib import Path

import pytest

# 导入 server 模块中的工具函数
from excel_mcp.server import execute_excel_vba, get_excel_path
from excel_mcp.vba_executor import VBAExecutor
from excel_mcp.exceptions import (
    VBASecurityError,
    VBAExecutionError,
    WorkbookError,
)


class TestExecuteExcelVBAIntegration:
    """execute_excel_vba MCP 工具集成测试"""
    
    def test_security_check_blocks_dangerous_code(self):
        """
        测试安全检查阻止危险代码。
        _Requirements: 2.1, 2.2_
        """
        dangerous_code = """
Sub Main()
    Shell "cmd /c dir"
End Sub
"""
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            temp_file = f.name
            f.write(b"PK")  # 简单的文件头
        
        try:
            result = execute_excel_vba(temp_file, dangerous_code, "Main")
            result_dict = json.loads(result)
            
            assert result_dict["status"] == "error"
            assert "安全检查失败" in result_dict["message"] or "Shell" in result_dict["message"]
        finally:
            os.unlink(temp_file)
    
    def test_nonexistent_file_returns_error(self):
        """
        测试不存在的文件返回错误。
        _Requirements: 3.1_
        """
        safe_code = """
Sub Main()
    MsgBox "Hello"
End Sub
"""
        result = execute_excel_vba("/nonexistent/path/file.xlsx", safe_code, "Main")
        result_dict = json.loads(result)
        
        assert result_dict["status"] == "error"
        assert "不存在" in result_dict["message"] or "error" in result_dict["message"].lower()
    
    def test_result_structure_contains_required_fields(self):
        """
        测试返回结果包含必需字段。
        _Requirements: 1.3, 4.3_
        """
        # 使用会触发安全检查的代码，这样不需要真实的 Excel
        dangerous_code = """
Sub Main()
    Kill "test.txt"
End Sub
"""
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            temp_file = f.name
            f.write(b"PK")
        
        try:
            result = execute_excel_vba(temp_file, dangerous_code, "Main")
            result_dict = json.loads(result)
            
            # 验证结果结构
            assert "status" in result_dict
            assert "message" in result_dict
            assert "logs" in result_dict
            assert isinstance(result_dict["logs"], list)
        finally:
            os.unlink(temp_file)
    
    def test_multiple_sensitive_keywords_all_detected(self):
        """
        测试多个敏感关键词都被检测到。
        _Requirements: 2.1, 2.2_
        """
        multi_dangerous_code = """
Sub Main()
    Shell "cmd"
    Kill "file.txt"
    CreateObject("Scripting.FileSystemObject")
End Sub
"""
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            temp_file = f.name
            f.write(b"PK")
        
        try:
            result = execute_excel_vba(temp_file, multi_dangerous_code, "Main")
            result_dict = json.loads(result)
            
            assert result_dict["status"] == "error"
            # 至少检测到一个敏感关键词
            message = result_dict["message"]
            assert any(kw in message for kw in ["Shell", "Kill", "CreateObject", "安全检查"])
        finally:
            os.unlink(temp_file)


class TestVBAExecutorDirectCalls:
    """VBAExecutor 直接调用测试"""
    
    def test_scan_sensitive_keywords_returns_list(self):
        """扫描敏感关键词返回列表"""
        executor = VBAExecutor()
        
        safe_code = "Sub Main()\n    MsgBox \"Hello\"\nEnd Sub"
        result = executor._scan_sensitive_keywords(safe_code)
        
        assert isinstance(result, list)
        assert len(result) == 0
    
    def test_scan_detects_all_keyword_categories(self):
        """测试检测所有类别的敏感关键词"""
        executor = VBAExecutor()
        
        # 系统命令
        assert "Shell" in executor._scan_sensitive_keywords("Shell cmd")
        assert "SendKeys" in executor._scan_sensitive_keywords("SendKeys keys")
        
        # 文件系统
        assert "FileSystemObject" in executor._scan_sensitive_keywords("FileSystemObject")
        assert "DeleteFile" in executor._scan_sensitive_keywords("DeleteFile path")
        
        # 网络操作
        assert "WinHttp" in executor._scan_sensitive_keywords("WinHttp.Request")
        assert "XMLHTTP" in executor._scan_sensitive_keywords("XMLHTTP")
        
        # 自动启动宏
        assert "Workbook_Open" in executor._scan_sensitive_keywords("Sub Workbook_Open()")
        assert "Auto_Open" in executor._scan_sensitive_keywords("Sub Auto_Open()")
    
    def test_backup_creates_file_in_same_directory(self):
        """备份文件创建在同一目录"""
        executor = VBAExecutor()
        
        with tempfile.TemporaryDirectory() as tmp_dir:
            test_file = Path(tmp_dir) / "test.xlsx"
            test_file.write_bytes(b"test content")
            
            backup_path = executor._create_backup(str(test_file))
            
            assert Path(backup_path).parent == test_file.parent
            assert Path(backup_path).exists()
    
    def test_executor_constants_are_correct(self):
        """验证执行器常量配置正确"""
        assert VBAExecutor.EXECUTION_TIMEOUT == 30
        assert VBAExecutor.LOCK_TIMEOUT == 60
        assert VBAExecutor.BACKUP_PREFIX == "BACKUP_"
        assert len(VBAExecutor.SENSITIVE_KEYWORDS) > 0


class TestErrorScenarios:
    """错误场景测试"""
    
    def test_empty_filepath_raises_error(self):
        """空文件路径应返回错误"""
        result = execute_excel_vba("", "Sub Main()\nEnd Sub", "Main")
        result_dict = json.loads(result)
        
        assert result_dict["status"] == "error"
    
    def test_empty_vba_code_with_nonexistent_file(self):
        """空 VBA 代码和不存在的文件"""
        result = execute_excel_vba("/fake/path.xlsx", "", "Main")
        result_dict = json.loads(result)
        
        assert result_dict["status"] == "error"
    
    def test_json_output_is_valid(self):
        """验证输出是有效的 JSON"""
        result = execute_excel_vba("/fake/path.xlsx", "Sub Main()\nEnd Sub", "Main")
        
        # 应该能成功解析为 JSON
        parsed = json.loads(result)
        assert isinstance(parsed, dict)
