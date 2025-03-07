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

#ifndef YXLSX_ABSTRACTSHEET_H
#define YXLSX_ABSTRACTSHEET_H

#include "abstractooxmlfile.h"

QT_BEGIN_NAMESPACE_YXLSX

enum class SheetType { kWorkSheet };

class AbstractSheet : public AbstractOOXmlFile {
public:
    inline const QString& GetSheetName() const { return sheet_name_; }
    inline SheetType GetSheetType() const { return sheet_type_; }
    inline int GetSheetId() const { return sheet_id_; }

    inline void SetSheetName(const QString& sheet_ame) { sheet_name_ = sheet_ame; }
    inline void SetSheetType(SheetType sheet_type = SheetType::kWorkSheet) { sheet_type_ = sheet_type; }

protected:
    AbstractSheet(const QString& sheet_name, int sheet_id, SheetType sheet_type = SheetType::kWorkSheet)
        : sheet_name_ { sheet_name }
        , sheet_id_ { sheet_id }
        , sheet_type_ { sheet_type }
    {
    }

    QString sheet_name_ {};
    int sheet_id_ {};
    SheetType sheet_type_ {};
};

QT_END_NAMESPACE_YXLSX
#endif // YXLSX_ABSTRACTSHEET_H
