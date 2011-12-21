FINDPATH
--------

FindPath is a console-based utility that locates a file using the standard
search path of the Windows operating system. Given just the name of the file,
FindPath will search for the file in the standard search locations and, if
found, display its full path. You can additionally have FindPath copy the
found path to the clipboard or directly open it in the Windows Explorer.

Although FindPath can be used for locating any type of file, it was primarily
designed for the purpose of locating executable binaries and DLLs along the
standard search path. It is particularly handy when you have more than one
version of a file and you want to determine which is being located by the
system based on the current environment settings.

FindPath is also capable of working with [application and assembly manifests]
introduced with Windows XP in support of [isolated applications and
side-by-side assemblies]. Internally, FindPath is just an executable around
the [SearchPath] API.

  [application and assembly manifests]: http://msdn.microsoft.com/library/en-us/sbscs/setup/manifests.asp
  [isolated applications and side-by-side assemblies]: http://msdn.microsoft.com/library/en-us/sbscs/setup/isolated_applications_and_side_by_side_assemblies_start_page.asp
  [SearchPath]: http://msdn2.microsoft.com/en-us/library/aa365527.aspx
