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

QT_BEGIN_NAMESPACE_YXLSX

// ECMA 376, 18.18.11. ST_CellType (Cell Type)
// https://ecma-international.org/publications-and-standards/standards/ecma-376/
enum class CellType { kBoolean, kDate, kNumber, kSharedString, kUnknown };

struct Cell {
    explicit Cell(const QVariant& value = {}, CellType type = CellType::kNumber)
        : type { type }
        , value { value }
    {
    }

    explicit Cell(const Cell* const cell)
        : type { cell ? cell->type : CellType::kNumber }
        , value { cell ? cell->value : QVariant {} }
    {
    }

    CellType type {};
    QVariant value {};
};

QT_END_NAMESPACE_YXLSX
#endif // YXLSX_CELL_H
