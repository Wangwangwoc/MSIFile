# MSIFile
Used to extract information from an MSI file。使用Ctypes重写WINAPI中部分读取MSI文件信息的API，方便调用。


# 描述
最近有Ptthon2.7提取MSI ICON的需求，使用MSIlib只能做到读取String信息，没有读取Stream信息，其实也就差一个函数，但是发现Ctypes写WinAPI只要不涉及复杂结构体的情况下也很简单。所以就重写了。


