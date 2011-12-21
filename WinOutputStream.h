#pragma once

// --------------------------------------------------------------------------
//  WinOutputStream
// --------------------------------------------------------------------------
//
//  This class was mainly designed to avoid linking in the C/C++ RTL by
//  relying on the underyling OS API for minimal console output support. 
//
//  NOTE: The goal of relying on the C/C++ RTL has not been achieved fully 
//  for the overall application yet so the CRT is still linked in.
//

class WinOutputStream
{
public:

    WinOutputStream(HANDLE consoleHandle) : m_consoleHandle(consoleHandle) 
    { 
        _ASSERT(consoleHandle); 
    }

    void Write(LPCTSTR text)
    {
        _ASSERT(text);

        DWORD bytesWritten;
        DWORD textSize = lstrlen(text) * sizeof(TCHAR);
        WriteFile(m_consoleHandle, text, textSize, &bytesWritten, NULL);
    }
    
    void Write(const unsigned long n)
    {
        TCHAR text[20];
        wsprintf(text, _T("%lu"), n);
        Write(text);
    }

    void Write(const int n)
    {
        TCHAR text[20];
        wsprintf(text, _T("%d"), n);
        Write(text);
    }

    void Write(const TCHAR ch)
    {
        DWORD bytesWritten;
        WriteFile(m_consoleHandle, &ch, 1, &bytesWritten, NULL);
    }

private:

    HANDLE m_consoleHandle;

    WinOutputStream(const WinOutputStream&);
    WinOutputStream& operator=(const WinOutputStream&);
};
