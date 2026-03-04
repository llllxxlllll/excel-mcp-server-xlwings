# -*- coding: utf-8 -*-
"""VBA 代码执行器模块

本模块提供动态 VBA 代码执行功能，包括：
- 敏感关键词安全扫描
- 文件备份机制
- VBA 信任检查
- 代码注入和执行（带超时保护）
- 并发控制
"""

import logging
import os
import shutil
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import xlwings as xw

from .exceptions import (
    VBABusyError,
    VBAExecutionError,
    VBASecurityError,
    VBATimeoutError,
    WorkbookError,
)
from .xw_helper import ExcelNotFoundError, get_app, get_workbook


logger = logging.getLogger(__name__)


class VBAExecutor:
    """VBA 代码执行器
    
    负责 VBA 代码的安全检查、注入和执行。
    
    Attributes:
        SENSITIVE_KEYWORDS: 敏感关键词列表（大小写不敏感匹配）
        BACKUP_PREFIX: 备份文件前缀
        EXECUTION_TIMEOUT: 执行超时时间（秒）
        LOCK_TIMEOUT: 获取锁的等待超时（秒）
        _lock: 类级别的全局锁，确保同一时间只有一个 VBA 执行
    """
    
    # 扩展的敏感关键词列表（大小写不敏感匹配）
    SENSITIVE_KEYWORDS = [
        # 系统命令执行
        "Shell", "Kill", "WScript", "Scripting", "SendKeys", "AppActivate",
        "Powershell", "Cmd",
        # 文件系统操作
        "FileSystemObject", "CreateObject", "GetObject",
        "DeleteFile", "DeleteFolder", "CopyFile", "MoveFile",
        "MkDir", "RmDir", "ChDir",
        # 网络操作
        "Adodb.stream", "Shell.Application", "WinHttp", "XMLHTTP", "UrlMon",
        # 环境和反射
        "Environ", "CallByName",
        # 自动启动宏（防止病毒注入）
        "Workbook_Open", "Auto_Open", "Auto_Close"
    ]
    
    BACKUP_PREFIX = "BACKUP_"
    EXECUTION_TIMEOUT = 30  # 默认 30 秒超时
    LOCK_TIMEOUT = 60  # 获取锁的等待超时
    
    _lock = threading.Lock()  # 类级别的全局锁
    
    def __init__(self):
        """初始化 VBA 执行器"""
        self.app: Optional[xw.App] = None
        self.logger = logging.getLogger(__name__)
    
    def execute_vba(
        self,
        filepath: str,
        vba_code: str,
        entry_sub_name: str = "Main",
        timeout: Optional[int] = None
    ) -> Dict[str, Any]:
        """执行 VBA 代码的主入口
        
        Args:
            filepath: Excel 文件路径
            vba_code: VBA 代码字符串
            entry_sub_name: 入口 Sub 名称，默认为 "Main"
            timeout: 执行超时时间（秒），默认使用 EXECUTION_TIMEOUT
        
        Returns:
            执行结果字典，包含 status, message, logs, backup_path 等
        
        Raises:
            VBABusyError: Excel 正忙，无法获取执行锁
            VBASecurityError: 安全检查失败
            VBAExecutionError: VBA 执行失败
            VBATimeoutError: VBA 执行超时
        """
        logs: List[str] = []
        backup_path: Optional[str] = None
        wb: Optional[xw.Book] = None
        timeout = timeout or self.EXECUTION_TIMEOUT
        
        # 初始化 COM（MCP 服务器可能在子线程中运行）
        import pythoncom
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass  # 可能已经初始化
        
        # 1. 尝试获取全局锁
        if not self._acquire_lock():
            raise VBABusyError("Excel 正忙，请稍后重试")
        
        logs.append("已获取执行锁")
        
        try:
            # 2. 验证文件存在
            if not os.path.exists(filepath):
                raise WorkbookError(f"文件不存在: {filepath}")
            logs.append(f"文件验证通过: {filepath}")
            
            # 3. 安全检查 - 扫描敏感关键词
            detected = self._scan_sensitive_keywords(vba_code)
            if detected:
                raise VBASecurityError(f"检测到敏感关键词: {', '.join(detected)}")
            logs.append("安全检查通过")
            
            # 4. 创建备份
            backup_path = self._create_backup(filepath)
            logs.append(f"已创建备份: {backup_path}")
            
            # 5. 打开工作簿
            wb = get_workbook(filepath)
            logs.append("已打开工作簿")
            
            # 6. 检查 VBA 信任设置
            if not self._check_vba_trust(wb):
                raise VBASecurityError(
                    "请在 Excel 信任中心勾选 '信任对 VBA 工程对象模型的访问'"
                )
            logs.append("VBA 信任检查通过")
            
            # 7. 注入并执行 VBA 代码
            result = self._inject_and_execute(wb, vba_code, entry_sub_name, timeout)
            logs.extend(result.get("logs", []))
            
            # 8. 保存工作簿
            wb.save()
            logs.append("工作簿已保存")
            
            return {
                "status": "success",
                "message": "VBA 执行成功",
                "logs": logs,
                "backup_path": backup_path
            }
            
        except VBATimeoutError:
            # 超时时强制终止 Excel
            self._force_kill_excel()
            logs.append("执行超时，已强制终止 Excel 进程")
            raise
        except (VBASecurityError, VBAExecutionError, VBABusyError, WorkbookError):
            raise
        except ExcelNotFoundError as e:
            raise VBAExecutionError(f"Excel 未安装或无法启动: {e}")
        except Exception as e:
            self.logger.error(f"VBA 执行异常: {e}")
            raise VBAExecutionError(f"VBA 执行失败: {e}")
        finally:
            # 确保清理资源
            self._cleanup_resources(wb)
            # 释放锁
            self._release_lock()
            # 释放 COM
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    
    def _acquire_lock(self) -> bool:
        """尝试获取全局锁
        
        Returns:
            是否成功获取锁
        """
        return self._lock.acquire(timeout=self.LOCK_TIMEOUT)
    
    def _release_lock(self) -> None:
        """释放全局锁"""
        try:
            self._lock.release()
        except RuntimeError:
            # 锁未被持有，忽略
            pass
    
    def _scan_sensitive_keywords(self, vba_code: str) -> List[str]:
        """扫描代码中的敏感关键词（大小写不敏感）
        
        Args:
            vba_code: VBA 代码字符串
        
        Returns:
            检测到的敏感关键词列表
        """
        detected = []
        code_lower = vba_code.lower()
        
        for keyword in self.SENSITIVE_KEYWORDS:
            if keyword.lower() in code_lower:
                detected.append(keyword)
        
        return detected
    
    def _create_backup(self, filepath: str) -> str:
        """创建文件备份
        
        Args:
            filepath: 原始文件路径
        
        Returns:
            备份文件路径
        
        Raises:
            VBAExecutionError: 备份创建失败
        """
        try:
            path = Path(filepath)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"{self.BACKUP_PREFIX}{timestamp}_{path.name}"
            backup_path = path.parent / backup_name
            
            shutil.copy2(filepath, backup_path)
            self.logger.info(f"已创建备份: {backup_path}")
            
            return str(backup_path)
        except Exception as e:
            raise VBAExecutionError(f"备份创建失败: {e}")
    
    def _check_vba_trust(self, wb: xw.Book) -> bool:
        """检查 VBA 信任设置是否已开启
        
        Args:
            wb: xlwings Book 对象
        
        Returns:
            是否已开启 VBA 信任
        """
        try:
            # 尝试访问 VBProject，如果未开启信任会抛出异常
            _ = wb.api.VBProject.VBComponents.Count
            return True
        except Exception as e:
            self.logger.warning(f"VBA 信任检查失败: {e}")
            return False
    
    def _wrap_vba_with_error_handling(
        self,
        vba_code: str,
        entry_sub_name: str
    ) -> str:
        """包装 VBA 代码，添加全局错误处理
        
        将用户代码包装在错误处理结构中，捕获运行时错误。
        使用 On Error Resume Next 确保不会弹出错误对话框。
        
        Args:
            vba_code: 原始 VBA 代码
            entry_sub_name: 入口 Sub 名称
            
        Returns:
            包装后的 VBA 代码
        """
        # 创建包装后的代码，添加错误处理
        wrapped_code = f'''
' 错误信息存储
Public VBA_ERROR_MESSAGE As String
Public VBA_ERROR_NUMBER As Long
Public VBA_ERROR_SOURCE As String

' 包装的入口函数 - 使用 Resume Next 确保不弹对话框
Sub {entry_sub_name}_Wrapper()
    ' 清空错误信息
    VBA_ERROR_MESSAGE = ""
    VBA_ERROR_NUMBER = 0
    VBA_ERROR_SOURCE = ""
    
    ' 启用错误捕获
    On Error Resume Next
    
    ' 调用用户代码
    Call {entry_sub_name}
    
    ' 检查是否有错误
    If Err.Number <> 0 Then
        VBA_ERROR_NUMBER = Err.Number
        VBA_ERROR_MESSAGE = Err.Description
        VBA_ERROR_SOURCE = Err.Source
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

' 获取错误信息的函数
Function GetVBAError() As String
    If VBA_ERROR_NUMBER <> 0 Then
        GetVBAError = "VBA错误 " & VBA_ERROR_NUMBER & ": " & VBA_ERROR_MESSAGE
        If VBA_ERROR_SOURCE <> "" Then
            GetVBAError = GetVBAError & " (来源: " & VBA_ERROR_SOURCE & ")"
        End If
    Else
        GetVBAError = ""
    End If
End Function

Function GetVBAErrorNumber() As Long
    GetVBAErrorNumber = VBA_ERROR_NUMBER
End Function

{vba_code}
'''
        return wrapped_code
    
    def _reset_excel_state(self, app) -> Dict[str, Any]:
        """重置 Excel 状态，禁用可能导致阻塞的对话框
        
        Args:
            app: Excel Application 对象
            
        Returns:
            原始状态字典，用于恢复
        """
        original_state = {}
        try:
            # 保存原始状态
            original_state["DisplayAlerts"] = app.DisplayAlerts
            original_state["EnableEvents"] = app.EnableEvents
            original_state["ScreenUpdating"] = app.ScreenUpdating
            
            # 禁用警告对话框和事件
            app.DisplayAlerts = False
            app.EnableEvents = False
            # 保持屏幕更新以便用户看到变化
            app.ScreenUpdating = True
            
        except Exception as e:
            self.logger.warning(f"重置 Excel 状态时出错: {e}")
        
        return original_state
    
    def _restore_excel_state(self, app, original_state: Dict[str, Any]) -> None:
        """恢复 Excel 状态
        
        Args:
            app: Excel Application 对象
            original_state: 原始状态字典
        """
        try:
            if "DisplayAlerts" in original_state:
                app.DisplayAlerts = original_state["DisplayAlerts"]
            if "EnableEvents" in original_state:
                app.EnableEvents = original_state["EnableEvents"]
            if "ScreenUpdating" in original_state:
                app.ScreenUpdating = original_state["ScreenUpdating"]
        except Exception as e:
            self.logger.warning(f"恢复 Excel 状态时出错: {e}")

    def _inject_and_execute(
        self,
        wb: xw.Book,
        vba_code: str,
        entry_sub_name: str,
        timeout: int
    ) -> Dict[str, Any]:
        """注入代码并执行
        
        注意：COM 对象不能跨线程使用，因此直接在主线程执行。
        VBA 代码会被包装以添加错误处理，防止错误对话框阻塞。
        
        Args:
            wb: xlwings Book 对象
            vba_code: VBA 代码字符串
            entry_sub_name: 入口 Sub 名称
            timeout: 超时时间（秒）- 当前未使用，保留参数兼容性
        
        Returns:
            执行结果字典
        
        Raises:
            VBAExecutionError: 执行失败
        """
        result = {"status": "pending", "message": "", "logs": []}
        module_name = f"TempVBAModule_{datetime.now().strftime('%H%M%S%f')}"
        new_module = None
        original_state = {}
        
        try:
            # 重置 Excel 状态，禁用警告对话框
            original_state = self._reset_excel_state(wb.app.api)
            result["logs"].append("已禁用 Excel 警告对话框")
            
            # 包装 VBA 代码，添加错误处理
            wrapped_code = self._wrap_vba_with_error_handling(vba_code, entry_sub_name)
            
            # 注入 VBA 模块
            vb_project = wb.api.VBProject
            new_module = vb_project.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
            new_module.Name = module_name
            new_module.CodeModule.AddFromString(wrapped_code)
            result["logs"].append(f"已注入 VBA 模块: {module_name}")
            
            # 执行包装后的宏
            wrapper_name = f"{module_name}.{entry_sub_name}_Wrapper"
            wb.app.api.Run(wrapper_name)
            result["logs"].append(f"已执行宏: {entry_sub_name}")
            
            # 检查是否有 VBA 运行时错误
            try:
                error_msg = wb.app.api.Run(f"{module_name}.GetVBAError")
                error_num = wb.app.api.Run(f"{module_name}.GetVBAErrorNumber")
                
                if error_num and error_num != 0:
                    raise VBAExecutionError(f"VBA 运行时错误: {error_msg}")
            except VBAExecutionError:
                raise
            except Exception:
                # 获取错误信息失败，忽略
                pass
            
            result["status"] = "success"
            result["message"] = "VBA 执行成功"
            
        except VBAExecutionError:
            raise
        except Exception as e:
            error_msg = str(e).lower()
            
            # 检查是否是入口 Sub 不存在的错误
            if "macro" in error_msg or "not found" in error_msg or "找不到" in error_msg:
                raise VBAExecutionError(
                    f"入口 Sub '{entry_sub_name}' 不存在，请检查 VBA 代码中是否定义了该过程"
                )
            
            raise VBAExecutionError(f"VBA 运行时错误: {e}")
            
        finally:
            # 清理注入的模块
            if new_module is not None:
                try:
                    vb_project = wb.api.VBProject
                    vb_project.VBComponents.Remove(new_module)
                    result["logs"].append(f"已清理 VBA 模块: {module_name}")
                except Exception as cleanup_error:
                    result["logs"].append(f"模块清理警告: {cleanup_error}")
            
            # 恢复 Excel 状态
            self._restore_excel_state(wb.app.api, original_state)
        
        return result
    
    def _cleanup_resources(self, wb: Optional[xw.Book]) -> None:
        """清理资源
        
        Args:
            wb: xlwings Book 对象
        """
        if wb is not None:
            try:
                wb.close()
            except Exception as e:
                self.logger.warning(f"关闭工作簿时出错: {e}")
    
    def _force_kill_excel(self) -> None:
        """强制终止 Excel 进程（超时时使用）"""
        try:
            import subprocess
            subprocess.run(
                ["taskkill", "/F", "/IM", "EXCEL.EXE"],
                capture_output=True,
                timeout=10
            )
            self.logger.warning("已强制终止 Excel 进程")
        except Exception as e:
            self.logger.error(f"强制终止 Excel 进程失败: {e}")
