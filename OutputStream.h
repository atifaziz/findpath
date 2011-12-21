#pragma once

//
//  These operators were mainly designed to avoid linking in the C/C++ RTL.
//  They provide the minimal support for expressing the "insertion" 
//  style of writing to streams. 
//
//  NOTE: The goal of relying on the C/C++ RTL has not been achieved fully 
//  for the overall application yet so the CRT is still linked in.
//

template<class T>
T& operator<<(T& t, LPCTSTR text)
{
    _ASSERT(text);
    t.Write(text);
    return t;
}

template<class T>
T& operator<<(T& t, const unsigned long n)
{
    t.Write(n);
    return t;
}

template<class T>
T& operator<<(T& t, const int n)
{
    t.Write(n);
    return t;
}

template<class T>
T& operator<<(T& t, const TCHAR c)
{
    t.Write(c);
    return t;
}
