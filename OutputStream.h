// FINDPATH - Locates a file using the Windows search path
// Copyright (C) 2002, Atif Aziz (http://www.raboof.com)
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
