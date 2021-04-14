#!/usr/bin/env python
# -- coding: utf-8 --
# @Author : wxy
# @File : MSIFile.py

import ctypes
msidll = ctypes.windll.msi

def storeHandle(func):
    """装饰器函数，用于操作存储句柄和指针"""
    def store(self, *args, **kwargs):
        Handle, pHandle = func(self, *args, **kwargs)
        self.__handle_list__.append(Handle)
        self.__phandle_list__.append(pHandle)
        return Handle, pHandle
    return store


class MsiFile:
    """
    参考 msiquery.h, msi.h 使用Ctypes重写操作MSI文件的WinAPI
    因需求原因，只涉及读取信息，关于写msi文件信息可参考pypi-msilib
    """

    def __init__(self,file_path,sql):

        #以下列表保存所有打开的句柄和对应指针
        self.__handle_list__ = []
        self.__phandle_list__ = []

        #根据微软文档，按步骤打开数据库执行SQL语句查询信息
        try:
            self.hDataBase, self.phDataBase = self.OpenDataBase(file_path)
            self.hView, self.phView = self.DatabaseOpenViewW(sql)
            self.hRecord, self.phRecord = self.Execute()
        except:
            raise
        # 必须先Fetch获取结果才能获取结果，对多行结果的，Fetch一次获取一行结果。

    def close(self):
        """对象使用结束调用close以关闭句柄"""
        for handle in self.__handle_list__:
            msidll.MsiCloseHandle(handle)

    def __del__(self):
        self.close()
    @storeHandle
    def OpenDataBase(self,file_path):
        hDataBase = ctypes.c_ulong(0)
        phDataBase = ctypes.pointer(hDataBase)
        file_path = ctypes.c_char_p(file_path)
        szPersist = ctypes.c_char_p("1") # MSI_READ_ONLY
        if msidll.MsiOpenDatabaseA(file_path, szPersist, phDataBase) != 0:
            msidll.MsiCloseHandle(hDataBase)
            print("打开数据库句柄失败，检查文件路径和文件是否是MSI格式")
            raise
        return hDataBase,phDataBase

    @storeHandle
    def DatabaseOpenViewW(self,sql):
        psql = ctypes.c_char_p(sql)
        hView = ctypes.c_ulong(0)
        phView = ctypes.pointer(hView)
        if msidll.MsiDatabaseOpenViewA(self.hDataBase,psql,phView) != 0:
            msidll.MsiCloseHandle(hView)
            print("打开View对象失败,检查SQL语句是否正确")
            raise
        return hView,phView

    @storeHandle
    def Execute(self):
        hRecord = ctypes.c_ulong(0)
        msidll.MsiViewExecute(self.hView,hRecord)
        phRecord = ctypes.pointer(hRecord)
        return hRecord, phRecord

    def Fetch(self):
        """执行Execute后必须执行Fetch才能获得数据"""
        msidll.MsiViewFetch(self.hView,self.phRecord)

    def RecordGetFieldCount(self):
        """返回条目数量，即数据表中的列数量"""
        counts  = msidll.MsiRecordGetFieldCount(self.hRecord)
        return counts

    def RecordGetString(self,prev_string_length=8,field=1):
        """读字符串"""
        field = ctypes.c_uint(field)
        buf = (ctypes.c_char * prev_string_length)()
        pbuf = ctypes.pointer(buf)
        length = ctypes.c_ulong(prev_string_length)
        plength  = ctypes.pointer(length)
        msidll.MsiRecordGetStringA(self.hRecord,field,pbuf,plength)
        return buf

    def ReadStream(self, prev_stream_size=2048, field=1):
        """读二进制数据比较有用，默认每次读2048大小"""
        field = ctypes.c_uint(field)
        buf = (ctypes.c_char * prev_stream_size)()
        pbuf = ctypes.pointer(buf)
        buf_size = ctypes.sizeof(buf)
        flag_size = prev_stream_size
        buf_size = ctypes.c_long(buf_size)
        pbuf_size = ctypes.pointer(buf_size)
        status = 0
        res = []
        while status == 0:
            status = msidll.MsiRecordReadStream(
            self.hRecord, field, pbuf, pbuf_size)
            res.append(buf.raw)
            if buf_size.value != flag_size:
                break
        data = "".join(res)
        return data

#以下几个例子用于读取MSI数据库中的数据
def getIconData(file_path):
    """
    example: sql = "select Data  From Icon"
    获取ICON数据，可能为PE格式和ICON格式。后续再判断使用
    参考：https://docs.microsoft.com/en-us/windows/win32/msi/icon-table
    """
    sql = "select Data  From Icon"
    myMsiFile = MsiFile(file_path, sql)
    myMsiFile.Fetch()    #目前已知ICON表只有一行数据，所以只Fetch一次，有错误再改
    if myMsiFile.RecordGetFieldCount() >= 1:
        buf = myMsiFile.ReadStream(prev_stream_size=2048, field=1)
        return buf
    else:
        return None

if __name__ == "__main__":
    file_path = 'python-2.7.18.msi'
    # example: sql = "select Data  From Icon" # 获取ICON数据
    data = getIconData(file_path)
