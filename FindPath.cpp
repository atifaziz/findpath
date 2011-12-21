// FINDPATH - Locates a file using the Windows search path
// Copyright (C) 2004, Atif Aziz (http://www.raboof.com)
//
// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

#include "stdafx.h"
#include "Exceptions.h"
#include "OutputStream.h"
#include "WinOutputStream.h"

//
// Libraries
//

#pragma comment(lib, "version")
#pragma comment(lib, "shlwapi")

//
// Local functions
//

static void ShowHelp();
static void ShowLogo();
static void CopyToClipboard(LPCTSTR text);
static void OpenContainingFolder(LPCTSTR path);
static int SplitString(LPTSTR text, TCHAR delimiter);
static void ExtractManifest(LPCTSTR path);
static BOOL CALLBACK EnumResourceNamesCallback(HMODULE moduleHandle, LPCTSTR type, LPTSTR name, LONG_PTR userParam);

//
// Macros
//

#ifndef DIM
#define DIM(a) (sizeof(a) / sizeof((a)[0]))
#endif

//
// Global variables
//

WinOutputStream cout(GetStdHandle(STD_OUTPUT_HANDLE));
WinOutputStream cerr(GetStdHandle(STD_ERROR_HANDLE));

//
// Activation context API
//
// These function pointers are for late-binding support of the activation 
// context API to prevent the application from implicitly importing them
// and therefore becoming bound to only those versions of the operating
// system that implement them.
//

namespace ActivationContext
{
    HANDLE (WINAPI * Create)(PACTCTX);
    BOOL   (WINAPI * Activate)(HANDLE, ULONG_PTR*);
    BOOL   (WINAPI * Deactivate)(DWORD, ULONG_PTR);
    void   (WINAPI * Release)(HANDLE);
};

//
// Redirectiong of activaton context API via macros.
//

#undef CreateActCtx
#define CreateActCtx     ActivationContext::Create
#define ActivateActCtx   ActivationContext::Activate
#define DeactivateActCtx ActivationContext::Deactivate
#define ReleaseActCtx    ActivationContext::Release

// --------------------------------------------------------------------------
//  CommandLineArguments<Handler>
// --------------------------------------------------------------------------

template<class Handler>
class CommandLineArguments : public Handler
{
public:

    CommandLineArguments() {}

    bool Parse(int argumentsLength, LPCTSTR arguments[])
    {
        for (int i = 1; i < argumentsLength; i++)
        {
            LPCTSTR argument = arguments[i];

            if (argument[0] == _T('-') || 
                argument[0] == _T('/'))
            {
                const bool isLastArgument = (i + 1) == argumentsLength ;
                LPCTSTR nextArgument = isLastArgument ? NULL : arguments[i + 1];

                if (!Handler::HandleOption(argument + 1, nextArgument))
                    return false;

                if (!nextArgument)
                    i++;
            }
            else
            {
                if (!Handler::HandleUnnamed(argument))
                    return false;
            }
        }

        return Handler::EndOfParse();
    }

private:

    CommandLineArguments(const CommandLineArguments&);
    CommandLineArguments& operator=(const CommandLineArguments&);
};

// --------------------------------------------------------------------------
//  CommandLineHandler
// --------------------------------------------------------------------------

class CommandLineHandler 
{
public:

    LPCTSTR m_fileName;
    bool m_copyToClipboard;
    bool m_showHelp;
    bool m_openContainingFolder;
    bool m_verbose;
    bool m_suppressLogo;
    LPCTSTR m_manifestFilePath;
    bool m_extractManifest;

    CommandLineHandler() : 
        m_fileName(NULL),
        m_copyToClipboard(false),
        m_showHelp(false),
        m_openContainingFolder(false),
        m_verbose(false),
        m_suppressLogo(false),
        m_manifestFilePath(NULL),
        m_extractManifest(false)
        {}

    bool HandleUnnamed(LPCTSTR unnamed)
    {
        _ASSERT(unnamed);
        m_fileName = unnamed;

        return true;
    }

    bool HandleOption(LPCTSTR option, LPCTSTR& argument)
    {
        _ASSERT(option);

        if (IsOption(option, _T("nologo")))
        {
            m_suppressLogo = true;
        }
        else if (IsOption(option, _T("xm")))
        {
            m_extractManifest = true;
        }
        else
        {
            switch (tolower(option[0]))
            {
                case '?' : m_showHelp = true; break;
                case 'c' : m_copyToClipboard = true; break;
                case 'o' : m_openContainingFolder = true; break;
                case 'v' : m_verbose = true; break;

                case 'm' : 
                {
                    if (argument == NULL)
                    {
                        cerr << _T("Missing manifest file name.\n");
                        return false;
                    }

                    m_manifestFilePath = argument;
                    argument = NULL;
                    break;
                }

                default  : 
                {
                    cerr << _T("Invalid option: ") << option << _T("\n");
                    return false;
                }
            }
        }

        return true;
    }

    bool EndOfParse()
    {
        if (!m_showHelp && !m_fileName)
        {
            cerr << _T("Missing file name.\n");
            return false;
        }

        return true;
    }

private:

    static bool IsOption(LPCTSTR test, LPCTSTR option)
    {
        return CSTR_EQUAL == CompareString(LOCALE_INVARIANT, 0, test, -1, option, -1);
    }

    CommandLineHandler(const CommandLineHandler&);
    CommandLineHandler& operator=(const CommandLineHandler&);
};

// --------------------------------------------------------------------------
//  main
// --------------------------------------------------------------------------

template<class P>
void GetProcAddress(HMODULE module, LPCTSTR procedureName, P& procedure)
{
    procedure = (P) GetProcAddress(module, procedureName);
}

int _tmain(int argsLength, LPCTSTR args[])
{
#ifdef _DEBUG
    
    _CrtSetDbgFlag(_CrtSetDbgFlag(_CRTDBG_REPORT_FLAG) | 
        _CRTDBG_CHECK_ALWAYS_DF | _CRTDBG_LEAK_CHECK_DF);

#endif

    int exitCode = 0;
    LPTSTR path = NULL;
    LPTSTR formattedPath = NULL;
    HANDLE activationContext = NULL;
    ULONG_PTR activationContextActivationCookie = 0;
    HMODULE kernelLibrary = NULL;

    try
    {
        //
        // Parse command-line arguments.
        //

        CommandLineArguments<CommandLineHandler> arguments;

        if (!arguments.Parse(argsLength, args))
        {
            TCHAR moduleFileName[MAX_PATH];
            GetModuleFileName(NULL, moduleFileName, DIM(moduleFileName));
            PathStripPath(moduleFileName);
            PathRemoveExtension(moduleFileName);

            cerr << _T("\nTry '") << moduleFileName << _T(" -?' for more help.\n");

            return -1;
        }

        if (!arguments.m_suppressLogo)
            ShowLogo();

        //
        // If showing usage, then also exit program.
        //

        if (arguments.m_showHelp)
        {
            ShowHelp();
            return 0;
        }

        if (arguments.m_manifestFilePath)
        {
            kernelLibrary = LoadLibrary(_T("kernel32.dll"));
            
            bool isAvailable = false;

            if (NULL != kernelLibrary)
            {
                #ifdef UNICODE
                    GetProcAddress(kernelLibrary, _T("CreateActCtxW"), ActivationContext::Create);
                #else
                    GetProcAddress(kernelLibrary, _T("CreateActCtxA"), ActivationContext::Create);
                #endif

                GetProcAddress(kernelLibrary, _T("ActivateActCtx"), ActivationContext::Activate);
                GetProcAddress(kernelLibrary, _T("DeactivateActCtx"), ActivationContext::Deactivate);
                GetProcAddress(kernelLibrary, _T("ReleaseActCtx"), ActivationContext::Release);

                isAvailable = true;
            }

            if (isAvailable)
            {
                ACTCTX activationContextSetup = { sizeof(activationContextSetup) };
                
                activationContextSetup.lpSource = arguments.m_manifestFilePath;
                activationContext = CreateActCtx(&activationContextSetup);

                if (INVALID_HANDLE_VALUE == activationContext)
                    SystemException::ThrowLast();

                if (!ActivateActCtx(activationContext, &activationContextActivationCookie))
                    SystemException::ThrowLast();
            }
        }

        //
        // Get the PATHEXT enivornment variable and create an array of
        // extensions out of it. This is used during search to test
        // against different extensions in case an extension was not
        // supplied with the filename.
        //

        LPCTSTR pathExtensions[31] = { 0 };
        const int maxPathExtensionCount = DIM(pathExtensions) - 1;

        DWORD pathExtLength = GetEnvironmentVariable(_T("PATHEXT"), NULL, 0);
        LPTSTR pathExt = static_cast<LPTSTR>(_alloca(pathExtLength * sizeof(pathExt[0])));
        GetEnvironmentVariable(_T("PATHEXT"), pathExt, pathExtLength);

        int pathExtensionCount = SplitString(pathExt, _T(';'));

        if (pathExtensionCount > maxPathExtensionCount)
            pathExtensionCount = maxPathExtensionCount;

        LPCTSTR pathExtension = pathExt;

        for (int i = 0; i < pathExtensionCount; i++)
        {
            pathExtensions[i] = pathExtension;
            pathExtension += lstrlen(pathExtension) + 1;
        }

        //
        // Repeat search until all extensions have been tried.
        //

        LPCTSTR extension = NULL;
        int extensionIndex = -1;

        DWORD pathLength;

        do
        {
            if (arguments.m_verbose)
            {
                cout << _T("Searching for ") << arguments.m_fileName;

                if (extension)
                    cout << extension;
                
                cout << _T('\n');
            }

            //
            // Get the search path, re-allocating receiving string
            // as needed.
            //

            pathLength = 0;
            DWORD pathCapacity;

            do
            {
                delete [] path;
                
                path = new TCHAR[pathLength];
                
                if (pathLength)
                    path[0] = 0;
                
                if (!path)
                    throw SystemException(ERROR_NOT_ENOUGH_MEMORY);

                pathCapacity = pathLength;

                LPTSTR filePart;

                pathLength = SearchPath(NULL, 
                    arguments.m_fileName, extension, 
                    pathCapacity, path, &filePart);
            }
            while (pathLength > pathCapacity);

            //
            // Did SearchPath fail? Find out why.
            //

            if (0 == pathLength)
            {
                delete [] path;
                path = NULL;

                //
                // If it is because the file was not found, then try the 
                // next extension. Otherwise throw an exception holding 
                // the last system error generated by SearchPath.
                //

                DWORD lastError = GetLastError();

                if (ERROR_FILE_NOT_FOUND == lastError)
                {
                    extension = pathExtensions[++extensionIndex];

                    //
                    // If at the end of the extensions array then just
                    // raise an exception stating the file was not found.
                    //

                    if (!extension)
                        throw SystemException(lastError);
                }
                else
                {
                    throw SystemException(lastError);
                }
            }
        }
        while (0 == pathLength);

        //
        // Quote the path if there is space in it, for long file paths.
        //

        if (StrChr(path, _T(' ')))
        {
            formattedPath = new TCHAR[pathLength + 3];

            if (!formattedPath)
                throw SystemException(ERROR_NOT_ENOUGH_MEMORY);

            formattedPath[0] = _T('"');
            ::lstrcpy(formattedPath + 1, path);
            formattedPath[pathLength + 1] = _T('"');
            formattedPath[pathLength + 2] = 0;
        }
        else
        {
            formattedPath = path;
        }

        //
        // Display the path.
        //

        cout << formattedPath << _T('\n');

        //
        // Copy to the clipboard if requested.
        //

        if (arguments.m_copyToClipboard)
            CopyToClipboard(formattedPath);

        //
        // Open the containing folder in Windows Explorer if requested.
        //

        if (arguments.m_openContainingFolder)
            OpenContainingFolder(path);

        if (arguments.m_extractManifest)
            ExtractManifest(path);
    }
    catch (SystemException& e)
    {
        //
        // Get and display the error message corresponding to exception
        // code.
        //

        LPTSTR message = NULL;

        FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER |
            FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
            NULL, e.GetCode(), 0, 
            reinterpret_cast<LPTSTR>(&message), 0, NULL);

        cerr << _T('\n') << e.GetCode() << _T(": ");

        if (message)
        {
            cerr << message << _T('\n');
            LocalFree(message);
        }
        else
        {
            cerr << _T("(No description available for this error code)\n");
        }

        exitCode = -1;
    }
    catch (ApplicationException& e)
    {
        cerr << e.GetMessage() << _T('\n');

        exitCode = -1;
    }

    if (formattedPath != path)
        delete [] formattedPath;

    delete [] path;

    if (activationContext)
    {
        if (activationContextActivationCookie)
            DeactivateActCtx(0, activationContextActivationCookie);

        ReleaseActCtx(activationContext);
    }

    if (NULL != kernelLibrary)
        FreeLibrary(kernelLibrary);


    return exitCode;
}

// --------------------------------------------------------------------------
//  ShowHelp
// --------------------------------------------------------------------------

void ShowHelp()
{
    //
    //
    //

    TCHAR applicationBinaryName[MAX_PATH];
    GetModuleFileName(NULL, applicationBinaryName, DIM(applicationBinaryName));
    PathStripPath(applicationBinaryName);
    PathRemoveExtension(applicationBinaryName);

    //
    // Show main usage, annotating it with system directories.
    //

    TCHAR systemPath[MAX_PATH];
    GetSystemDirectory(systemPath, DIM(systemPath));

    TCHAR windowsPath[MAX_PATH];
    GetWindowsDirectory(windowsPath, DIM(windowsPath));

    cout << _T("Usage: ") << applicationBinaryName 
         << _T(" [-c] [-m <manifest>] [-nologo] [-o] [-v] [-xm] [-?]\n")
         << _T("       <filename>\n\n")
         << _T("Searches for the specified file in the following directories,\n")
         << _T("in the following sequence:\n\n")
         << _T("1. The directory from which the application loaded.\n")
         << _T("2. The current directory.\n")
         << _T("3. The Windows system directory (") << systemPath << _T(").\n")
         << _T("4. Windows NT/2000/XP: The 16-bit Windows system directory.\n")
         << _T("   The directory is ") << windowsPath << _T("\\SYSTEM.\n")
         << _T("5. The Windows directory (") << windowsPath << _T(").\n")
         << _T("6. The directories that are listed in the PATH environment variable.\n\n");

    //
    // Break up the directories in the PATH and display them individually.
    //

    DWORD environmentPathLength = GetEnvironmentVariable(_T("PATH"), NULL, 0);

    LPTSTR environmentPath = new TCHAR[environmentPathLength];
    
    if (!environmentPath)
        throw SystemException(ERROR_NOT_ENOUGH_MEMORY);

    GetEnvironmentVariable(_T("PATH"), environmentPath, environmentPathLength);

    cout << _T("PATH contains the following directories:\n\n");

    int pathCount = SplitString(environmentPath, _T(';'));
    LPCTSTR path = environmentPath;

    for (int pathIndex = 0; pathIndex < pathCount; pathIndex++)
    {
        cout << _T("    ") << path << _T('\n');
        path += lstrlen(path) + 1;
    }

    delete [] environmentPath;

    //
    // Display PATHEXT realted help.
    //

    DWORD pathExtLength = GetEnvironmentVariable(_T("PATHEXT"), NULL, 0);
    LPTSTR pathExt = static_cast<LPTSTR>(_alloca(pathExtLength * sizeof(pathExt[0])));
    GetEnvironmentVariable(_T("PATHEXT"), pathExt, pathExtLength);

    cout << _T("\nIf the file indicated in <filename> is not found then a search\n")
            _T("is conducted with the extensions from PATHEXT appended to\n")
            _T("<filename> each time, where:\n")
            _T("\n    PATHEXT = ") << pathExt <<
            _T("\n\nThe left-to-right order of extensions in PATHEXT is significant.\n\n");

    //
    // Break up the extension in PATHEXT and display textual file type
    // information about each.
    //

    int pathExtensionCount = SplitString(pathExt, _T(';'));

    if (pathExtensionCount > 0)
    {
        cout << _T("Extensions in PATHEXT represent the following file types:\n\n");

        LPCTSTR pathExtension = pathExt;

        for (int i = 0; i < pathExtensionCount; i++)
        {
            SHFILEINFO fileInfo = { 0 };
                
            SHGetFileInfo(pathExtension, FILE_ATTRIBUTE_NORMAL, &fileInfo, 
                sizeof(fileInfo), SHGFI_USEFILEATTRIBUTES | SHGFI_TYPENAME);

            cout << _T("    ") 
                 << (pathExtension + 1) << _T(" = ") << fileInfo.szTypeName 
                 << _T('\n');

            pathExtension += lstrlen(pathExtension) + 1;
        }

        cout << _T('\n');
    }

    //
    // Describe the options.
    //

    cout << _T("Options:\n\n")
            _T("c      - Copy path to the clipboard.\n")
            _T("m      - Search using dependencies in <manfiest>.\n")
            _T("nologo - Suppress logo.\n")
            _T("o      - Open containing folder in Windows Explorer.\n")
            _T("v      - Verbose mode.\n")
            _T("xm     - Extract manifest from PE image.\n")
            _T("?      - Show this help.\n");
}

// --------------------------------------------------------------------------
//  ShowLogo
// --------------------------------------------------------------------------

void ShowLogo()
{
    //
    // Display the application title.
    //

    cout << _T("FindPath Utility\n");

    //
    // Get the version number from VERSIONINFO resource of this module.
    // If the version number cannot be retrieved successfully, then
    // it is simply not displayed. No error messages are displayed or
    // propagated since this is a non-critical function.
    //

    union 
    {
        ULARGE_INTEGER Number;
        
        struct
        {
            WORD QFE;
            WORD Build;
            WORD Minor;
            WORD Major;
        };
    } 
    version;

    TCHAR applicationBinaryName[MAX_PATH];
    GetModuleFileName(NULL, applicationBinaryName, DIM(applicationBinaryName));
    
    DWORD reservedHandle;
    DWORD verInfoSize = GetFileVersionInfoSize(applicationBinaryName, 
        &reservedHandle);

    if (0 != verInfoSize)
    {
        void* verInfo = new BYTE[verInfoSize];

        if (GetFileVersionInfo(applicationBinaryName, reservedHandle,
                verInfoSize, verInfo))
        {
            UINT fixedFileInfoSize;
            VS_FIXEDFILEINFO* fixedFileInfo;
            
            if (VerQueryValue(verInfo, _T("\\"), 
                    reinterpret_cast<LPVOID*>(&fixedFileInfo), 
                    &fixedFileInfoSize))
            {
                version.Number.HighPart = fixedFileInfo->dwFileVersionMS;
                version.Number.LowPart = fixedFileInfo->dwFileVersionLS;

                cout << _T("Version ")
                     << version.Major << _T('.')
                     << version.Minor << _T('.')
                     << version.Build << _T('.')
                     << version.QFE << _T(", ");
            }
        }

        delete [] verInfo;
    }

    //
    // Display credits.
    //

    cout << _T("Written by Atif Aziz\n\n");
}

// --------------------------------------------------------------------------
//  CopyToClipboard
// --------------------------------------------------------------------------

void CopyToClipboard(LPCTSTR text)
{
    if (!OpenClipboard(NULL))
        SystemException::ThrowLast();

    HGLOBAL dataHandle = NULL;

    try
    {
        dataHandle = GlobalAlloc(GMEM_MOVEABLE, 
            sizeof(TCHAR) * lstrlen(text) + 1);

        if (!dataHandle)
            SystemException::ThrowLast();

        LPTSTR clipboardPath = static_cast<LPTSTR>(GlobalLock(dataHandle));

        if (!clipboardPath)
            SystemException::ThrowLast();

        lstrcpy(clipboardPath, text);

        GlobalUnlock(dataHandle);

        if (!EmptyClipboard())
            SystemException::ThrowLast();

        if (!SetClipboardData(CF_TEXT, dataHandle))
            SystemException::ThrowLast();

        CloseClipboard();
    }
    catch (...)
    {
        if (dataHandle)
            GlobalFree(dataHandle);

        CloseClipboard();

        throw;
    }
}

// --------------------------------------------------------------------------
//  OpenContainingFolder
// --------------------------------------------------------------------------

void OpenContainingFolder(LPCTSTR path)
{
    _ASSERT(path);

    LPTSTR parameters = new TCHAR[lstrlen(path) + 100];
    
    if (!parameters)
        throw SystemException(ERROR_NOT_ENOUGH_MEMORY);

    wsprintf(parameters, _T("/select,%s"), path);

    HINSTANCE instance = ShellExecute(NULL, _T("open"), 
        _T("explorer.exe"), parameters, NULL, SW_SHOWNORMAL);

    delete [] parameters;

    if (instance <= reinterpret_cast<HINSTANCE>(32))
    {
        struct
        {
            UINT_PTR code;
            LPCTSTR message;
        }
        static errorTable[] =
        {
            { 0, _T("The operating system is out of memory or resources.") },
            { ERROR_FILE_NOT_FOUND, _T("The specified file was not found.") },
            { ERROR_PATH_NOT_FOUND, _T("The specified path was not found.") },
            { ERROR_BAD_FORMAT, _T("The .exe file is invalid (non-Win32® .exe or error in .exe image).") },
            { SE_ERR_ACCESSDENIED, _T("The operating system denied access to the specified file. ") },
            { SE_ERR_ASSOCINCOMPLETE, _T("The file name association is incomplete or invalid.") },
            { SE_ERR_DDEBUSY, _T("The DDE transaction could not be completed because other DDE transactions were being processed.") },
            { SE_ERR_DDEFAIL, _T("The DDE transaction failed.") },
            { SE_ERR_DDETIMEOUT, _T("The DDE transaction could not be completed because the request timed out.") },
            { SE_ERR_DLLNOTFOUND, _T("The specified dynamic-link library was not found. ") },
            { SE_ERR_FNF, _T("The specified file was not found. ") },
            { SE_ERR_NOASSOC, _T("There is no application associated with the given file name extension.") },
            { SE_ERR_OOM, _T("There was not enough memory to complete the operation.") },
            { SE_ERR_PNF, _T("The specified path was not found.") },
            { SE_ERR_SHARE, _T("A sharing violation occurred.") },
        };

        UINT_PTR code = reinterpret_cast<UINT_PTR>(instance);

        for (int i = 0; i < DIM(errorTable); i++)
        {
            if (code == errorTable[i].code)
                throw ApplicationException(errorTable[i].message);
        }

        throw ApplicationException(_T("Windows shell could not open the containing folder."));
    }
}

// --------------------------------------------------------------------------
//  ExtractManifest
// --------------------------------------------------------------------------

void ExtractManifest(LPCTSTR path)
{
    _ASSERT(path);

    //
    // Assume the worst. This method won't succeed. If we manage to reach
    // the end, then we'll indicate success.
    //

    bool succeeded = false;

    //
    // Load the PE image as a data file. This is good enough to work
    // with most of the resource reading API.
    //

    HMODULE module = LoadLibraryEx(path, NULL, 
        DONT_RESOLVE_DLL_REFERENCES | LOAD_LIBRARY_AS_DATAFILE);

    //
    // Enumerate the RT_MANIFEST resources and pick the first one.
    // This actual picking is done in the callback.
    //

    LPCTSTR resourceName = NULL;

    if (module)
    {
        EnumResourceNames(module, RT_MANIFEST, 
            EnumResourceNamesCallback, 
            reinterpret_cast<ULONG_PTR>(&resourceName));
    }

    //
    // Using the resource name of the first manifest, ask the system
    // to find it.
    //

    HRSRC resourceHandle = NULL;

    if (resourceName)
        resourceHandle = FindResource(module, resourceName, RT_MANIFEST);

    //
    // Load the manifest.
    //

    HGLOBAL resource = NULL;

    if (resourceHandle)
        resource = LoadResource(module, resourceHandle);

    //
    // Find out the size of the manifest.
    //

    DWORD manifestSize = 0;
    
    if (resource)
        manifestSize = SizeofResource(module, resourceHandle);

    //
    // Get the actual manifest data.
    //

    LPVOID manifest = NULL;

    if (manifestSize)
        manifest = LockResource(resource);

    //
    // Create a file to write out the manifest to.
    //

    HANDLE file = NULL;

    TCHAR manifestFileName[MAX_PATH];

    if (manifest)
    {
        StrCpy(manifestFileName, path);
        PathStripPath(manifestFileName);
        StrCat(manifestFileName, _T(".manifest"));
        
        file = CreateFile(manifestFileName, GENERIC_WRITE, 0, NULL, 
            CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
    }

    //
    // Write out the manifest!
    //

    if (INVALID_HANDLE_VALUE != file)
    {
        DWORD bytesWritten;
        succeeded = WriteFile(file, manifest, manifestSize, &bytesWritten, NULL) != FALSE;

        _ASSERT(bytesWritten == manifestSize);
    }

    //
    // Clean-up.
    //

    if (module)
        FreeLibrary(module);

    if (INVALID_HANDLE_VALUE != file)
        CloseHandle(file);

    //
    // If all worked out then write a message indicating that the manifest 
    // was successfully extracted. Otherwise throw an exception using the
    // last error indicated by the system.
    //

    if (succeeded)
    {
        cout << _T("Manifest extracted to: ") << manifestFileName << _T('\n');
    }
    else
    {
        SystemException::ThrowLast();
    }
}

// --------------------------------------------------------------------------
//  EnumResourceNamesCallback
// --------------------------------------------------------------------------

BOOL CALLBACK EnumResourceNamesCallback(HMODULE /* module */,
    LPCTSTR /* type */, LPTSTR name, LONG_PTR userParam)
{
    LPCTSTR* outName = reinterpret_cast<LPCTSTR*>(userParam); 
    _ASSERT(outName);

    *outName = name;

    //
    // Only interested in the first callback, so always return FALSE.
    //

    return FALSE;
}

// --------------------------------------------------------------------------
//  SplitString
// --------------------------------------------------------------------------

int SplitString(LPTSTR text, TCHAR delimiter)
{
    _ASSERT(text);

    int count = 0;
    TCHAR delimiterString[] = { delimiter, 0 };

    for (LPTSTR textDelimiter = StrPBrk(text, delimiterString); textDelimiter; 
         textDelimiter = StrPBrk(textDelimiter + 1, delimiterString))
    {
        *textDelimiter = 0;
        count++;
    }

    return count + 1;
}
