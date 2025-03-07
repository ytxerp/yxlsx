/*
 * MIT License
 *
 * Copyright (c) 2024 YTX
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

#ifndef YXLSX_SHEETFORMATPROPS_H
#define YXLSX_SHEETFORMATPROPS_H

#include "namespace.h"

QT_BEGIN_NAMESPACE_YXLSX

// ECMA-376 Part1 18.3.1.81
// https://ecma-international.org/publications-and-standards/standards/ecma-376/
struct SheetFormatProps {
    SheetFormatProps(double default_col_width = 8.430f, double default_row_height = 15)
        : default_col_width(default_col_width)
        , default_row_height(default_row_height)
    {
    }

    double default_col_width {};
    double default_row_height {};
};

QT_END_NAMESPACE_YXLSX

#endif // YXLSX_SHEETFORMATPROPS_H
