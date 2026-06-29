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

#ifndef YXLSX_CELL_H
#define YXLSX_CELL_H

#include <QVariant>

#include "namespace.h"

YXLSX_BEGIN_NAMESPACE

// ECMA 376, 18.18.11. ST_CellType (Cell Type)
// https://ecma-international.org/publications-and-standards/standards/ecma-376/
enum class CellType { kEmpty, kBoolean, kDateTime, kNumber, kSharedString, kInlineString, kError };

enum class StringType { kSharedString, kInlineString };

// Cell is a lightweight value container.
// - No formulas
// - No formatting
// - Strings may be stored as shared strings or inline strings.
// - DateTime represents ISO 8601 date cells (t="d").
struct Cell final {
    Cell() = default;

    explicit Cell(const QVariant& value, CellType type)
        : type(type)
        , value(value)
    {
    }

    CellType type {};
    QVariant value {};
};

YXLSX_END_NAMESPACE

Q_DECLARE_TYPEINFO(yxlsx::Cell, Q_MOVABLE_TYPE);

#endif // YXLSX_CELL_H
