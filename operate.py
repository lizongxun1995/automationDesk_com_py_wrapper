from typing import Union, Any, Optional

import pywintypes
import win32com.client
import win32api
import time
import pathlib
import enum
from collections import namedtuple
import win32com.client.CLSIDToClass, pythoncom, pywintypes
from pywintypes import IID

# class ApplicationEvent(TAMAutomation._IADApplicationEvents):
#     def __init__(self, executor):
#         super().__init__(executor)
#
#     def OnProjectActivate(self, ActivatedProject):
#         print('----------------')

class _IADApplicationEvents:
    '_IApplicationEvents Interface'
    CLSID = CLSID_Sink = IID('{53B5562D-EBD2-42C5-8EAE-A3D771EF6C57}')
    coclass_clsid = IID('{38F6B8E9-DF29-304B-8B8F-03BE5CEBCF2F}')
    _public_methods_ = []  # For COM Server support
    _dispid_to_func_ = {
        5001: "OnError",
        5005: "OnProjectActivate",
        5010: "OnProjectClose",
        5011: "OnProjectClosed",
        5006: "OnProjectCreate",
        5003: "OnProjectCreated",
        5007: "OnProjectOpen",
        5004: "OnProjectOpened",
        5008: "OnProjectSave",
        5009: "OnProjectSaved",
        5002: "OnWrite",
    }
    defaultNamedNotOptArg = pythoncom.Empty
    def __init__(self, oobj=None):
        if oobj is None:
            self._olecp = None
        else:
            import win32com.server.util
            from win32com.server.policy import EventHandlerPolicy
            cpc = oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
            cp = cpc.FindConnectionPoint(self.CLSID_Sink)
            cookie = cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
            self._olecp, self._olecp_cookie = cp, cookie

    def __del__(self):
        try:
            self.close()
        except pythoncom.com_error:
            pass

    def close(self):
        if self._olecp is not None:
            cp, cookie, self._olecp, self._olecp_cookie = self._olecp, self._olecp_cookie, None, None
            cp.Unadvise(cookie)

    def _query_interface_(self, iid):
        import win32com.server.util
        if iid == self.CLSID_Sink: return win32com.server.util.wrap(self)

    def OnProjectClose(self, ClosingProject=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
        print('----------------------')

class AutomationDesk:
    class DataType(enum.Enum):
        STRING = 'String'
        INT = 'Int'
        FLOAT = 'Float'
        FILE = 'File'
        DATA_CONTAINER = 'DataContainer'

    Data = namedtuple('Data', ['data_name', 'data_value', 'data_type', 'data_path'])

    def __init__(self, project_path: Union[pathlib.Path, str]):
        # Create the COM server
        self._automation_desk = win32com.client.Dispatch("AutomationDesk.TAM")
        # Show the user interface of AutomationDesk
        self._automation_desk.Visible = True
        self._proj_file = None
        self._proj_obj = None
        self._libs_obj = self._automation_desk.Libraries
        self._main_libs_obj = self._libs_obj.Item("Main Library")
        # Get the Standard library
        self._std_lib = self._libs_obj.Item("Standard")
        _IADApplicationEvents(self._automation_desk)
        project_path = pathlib.Path(project_path)
        # 判断是否为adpx文件
        if project_path.suffix != '.adpx':
            raise RuntimeError('不是正确的adpx项目')
        # 判断项目是否存在
        if not project_path.exists():
            # 创建项目
            self._proj_file = self._automation_desk.Projects
            self._proj_obj = self._proj_file.Create(project_path, "Standard Project", 1)
        else:
            # 打开项目
            self._proj_file = self._automation_desk.Projects

            self._proj_obj = self._proj_file.ImportProject(project_path, 1)

    def create_seq(self, seq_name: str, *positions: str):
        self.create_folder(*positions)
        sequence_templ = self._std_lib.SubBlocks.Item("Sequence")
        parent_folder = self._proj_obj
        for position in positions:
            parent_folder = parent_folder.SubBlocks.Item(position)
        try:
            seq_obj = parent_folder.SubBlocks.Create(sequence_templ)
            seq_obj.Name = seq_name
        except Exception:
            ...

    def create_folder(self, *positions: str) -> Any:
        folder_templ = self._std_lib.SubBlocks.Item("Folder")
        parent_folder = self._proj_obj
        for position in positions:
            if position[0].isdigit():
                raise RuntimeError('文件夹不能以数字开头')
            try:
                folder_obj = parent_folder.SubBlocks.Item(position)
                parent_folder = folder_obj
            except Exception:
                # 文件夹不存在
                folder_obj = parent_folder.SubBlocks.Create(folder_templ)
                folder_obj.Name = position
                parent_folder = folder_obj
        return self.get_object(*positions)

    def get_object(self, *positions: str) -> Any:
        parent_folder = self._proj_obj
        for position in positions:
            try:
                parent_folder = parent_folder.SubBlocks.Item(position)
            except:
                parent_folder = parent_folder.DataObjects.Item(position)

        return parent_folder

    def create_data(self, name: str, data_type: 'AutomationDesk.DataType', value: Optional[Any] = None,
                    *positions: str) -> Data:

        folder = self.get_object(*positions)
        data_obj_tmpl = self._main_libs_obj.DataObjects.Item(data_type.value)

        try:
            data_obj = folder.DataObjects.Create(data_obj_tmpl)
        except Exception:
            data_obj = folder.ChildDataObjects.Create(data_obj_tmpl)
        try:
            data_obj.Name = name
        except:
            ...
        if data_type is not AutomationDesk.DataType.DATA_CONTAINER:
            data_obj.Value = value
        return self.Data(value, value, data_type, positions)

    def exit(self):
        self._proj_obj.Save()
        self._proj_obj.Close()


if __name__ == '__main__':
    ad = AutomationDesk(r'C:\Users\Public\Documents\FTPSERVER\dSPACE\Project\Project.adpx')
    ad.create_seq('SEQ_006', 'f2', 'f3')
    ad.create_folder('f1', 'f2')
    ad.create_data('DataContainer1', AutomationDesk.DataType.DATA_CONTAINER, '12', 'f1', 'f2')
    ad.create_data('string2', AutomationDesk.DataType.STRING, '12', 'f1', 'f2', 'DataContainer1')
    ad.create_data('string3', AutomationDesk.DataType.STRING, '11', 'f1', 'f2', 'DataContainer1')
    ad.create_data('string4', AutomationDesk.DataType.STRING, '10', 'f1', 'f2', 'DataContainer1')
    ad.exit()
