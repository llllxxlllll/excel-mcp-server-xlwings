# -*- coding: utf-8 -*-
"""VBA 异常类单元测试

验证 VBA 相关异常类的继承关系和消息格式。
"""

import pytest
from excel_mcp.exceptions import (
    ExcelMCPError,
    VBAExecutionError,
    VBASecurityError,
    VBATimeoutError,
    VBABusyError,
)


class TestVBAExceptionHierarchy:
    """测试 VBA 异常类的继承关系"""

    def test_vba_execution_error_inherits_from_excel_mcp_error(self):
        """VBAExecutionError 应继承自 ExcelMCPError"""
        assert issubclass(VBAExecutionError, ExcelMCPError)

    def test_vba_security_error_inherits_from_vba_execution_error(self):
        """VBASecurityError 应继承自 VBAExecutionError"""
        assert issubclass(VBASecurityError, VBAExecutionError)
        assert issubclass(VBASecurityError, ExcelMCPError)

    def test_vba_timeout_error_inherits_from_vba_execution_error(self):
        """VBATimeoutError 应继承自 VBAExecutionError"""
        assert issubclass(VBATimeoutError, VBAExecutionError)
        assert issubclass(VBATimeoutError, ExcelMCPError)

    def test_vba_busy_error_inherits_from_vba_execution_error(self):
        """VBABusyError 应继承自 VBAExecutionError"""
        assert issubclass(VBABusyError, VBAExecutionError)
        assert issubclass(VBABusyError, ExcelMCPError)


class TestVBAExceptionMessages:
    """测试 VBA 异常类的消息格式"""

    def test_vba_execution_error_message(self):
        """VBAExecutionError 应正确保存错误消息"""
        msg = "VBA 执行失败"
        error = VBAExecutionError(msg)
        assert str(error) == msg

    def test_vba_security_error_message(self):
        """VBASecurityError 应正确保存错误消息"""
        msg = "检测到敏感关键词: Shell, Kill"
        error = VBASecurityError(msg)
        assert str(error) == msg

    def test_vba_timeout_error_message(self):
        """VBATimeoutError 应正确保存错误消息"""
        msg = "VBA 执行超时（30秒）"
        error = VBATimeoutError(msg)
        assert str(error) == msg

    def test_vba_busy_error_message(self):
        """VBABusyError 应正确保存错误消息"""
        msg = "Excel 正忙，请稍后重试"
        error = VBABusyError(msg)
        assert str(error) == msg


class TestVBAExceptionRaising:
    """测试 VBA 异常类的抛出和捕获"""

    def test_catch_vba_execution_error_as_excel_mcp_error(self):
        """应能用 ExcelMCPError 捕获 VBAExecutionError"""
        with pytest.raises(ExcelMCPError):
            raise VBAExecutionError("测试错误")

    def test_catch_vba_security_error_as_vba_execution_error(self):
        """应能用 VBAExecutionError 捕获 VBASecurityError"""
        with pytest.raises(VBAExecutionError):
            raise VBASecurityError("安全检查失败")

    def test_catch_vba_timeout_error_as_vba_execution_error(self):
        """应能用 VBAExecutionError 捕获 VBATimeoutError"""
        with pytest.raises(VBAExecutionError):
            raise VBATimeoutError("执行超时")

    def test_catch_vba_busy_error_as_vba_execution_error(self):
        """应能用 VBAExecutionError 捕获 VBABusyError"""
        with pytest.raises(VBAExecutionError):
            raise VBABusyError("Excel 正忙")
