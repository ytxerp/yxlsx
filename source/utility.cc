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

#include "utility.h"

#include <QDebug>
#include <QRegularExpression>

YXLSX_BEGIN_NAMESPACE

QStringList Utility::SplitPath(const QString& path)
{
    if (path.isEmpty())
        return { QStringLiteral("."), {} };

    const auto index { qMax(path.lastIndexOf(QLatin1Char('/')), path.lastIndexOf(QLatin1Char('\\'))) };

    if (index == -1)
        return { QStringLiteral("."), path };

    return { path.left(index), path.mid(index + 1) };
}

/*
 * Return the .rel file path based on filePath
 */
QString Utility::GetRelFilePath(const QString& file_path)
{
    if (file_path.isEmpty())
        return QString();

    const auto index { qMax(file_path.lastIndexOf(QLatin1Char('/')), file_path.lastIndexOf(QLatin1Char('\\'))) };

    if (index == -1) {
        return QStringLiteral("_rels/%1.rels").arg(file_path);
    }

    return QStringLiteral("%1/_rels/%2.rels").arg(file_path.left(index), file_path.mid(index + 1));
}

/*
  Creates a valid and unique sheet name.
    - Minimum length: 1
    - Maximum length: 31
    - Invalid characters (/ \ ? * ] [ :) are replaced by a single space (' ').
    - Sheet names must not begin or end with an apostrophe ('').
    - Ensures the name is unique by appending a number if necessary.
    - Generates a default name ("Sheet <N>") if the proposed name is empty or invalid.
 */
QString Utility::GenerateSheetName(const QStringList& sheet_names, const QString& name_proposal, int& last_sheet_index)
{
    QString ret { name_proposal };

    // If name is empty, generate a default name
    if (ret.isEmpty()) {
        do {
            ++last_sheet_index;
            ret = QStringLiteral("Sheet %1").arg(last_sheet_index);
        } while (sheet_names.contains(ret));

        return ret;
    }

    // Unescape sheet name if it starts and ends with an apostrophe
    if (ret.length() >= 3 && ret.startsWith(QLatin1Char('\'')) && ret.endsWith(QLatin1Char('\''))) {
        ret = UnescapeSheetName(ret);
    }

    // Replace invalid characters with a space
    static const QRegularExpression invalid_chars(QStringLiteral("[/\\\\?*\\][:]"));
    ret.replace(invalid_chars, QStringLiteral(" "));

    // Ensure the name doesn't start or end with an apostrophe
    if (ret.startsWith(QLatin1Char('\'')))
        ret[0] = QLatin1Char(' ');
    if (ret.endsWith(QLatin1Char('\'')))
        ret[ret.size() - 1] = QLatin1Char(' ');

    // Truncate to the maximum allowed length
    if (ret.size() >= 32)
        ret = ret.left(31);

    // Ensure uniqueness by appending a number if necessary
    QString unique_name = ret;
    int counter = 1;
    while (sheet_names.contains(unique_name)) {
        QString suffix = QStringLiteral(" (%1)").arg(counter++);
        if (ret.size() + suffix.size() >= 32)
            unique_name = ret.left(31 - suffix.size()) + suffix;
        else
            unique_name = ret + suffix;
    }

    return unique_name;
}

QString Utility::UnescapeSheetName(const QString& sheetName)
{
    if (sheetName.length() <= 2 || !sheetName.startsWith(QLatin1Char('\'')) || !sheetName.endsWith(QLatin1Char('\''))) {
        qWarning("Invalid sheet name format: '%s'", qPrintable(sheetName));
        return QString();
    }

    QString name { sheetName.mid(1, sheetName.length() - 2) };
    name.replace(QStringLiteral("''"), QStringLiteral("'"));
    return name;
}

/*
 * Check whether the string `s` starts or ends with a space or whitespace character.
 */
bool Utility::IsSpacePreserveNeeded(const QString& s)
{
    // Early exit if the string is empty
    if (s.isEmpty())
        return false;

    // Check if the first or last character is a whitespace character
    return s.front().isSpace() || s.back().isSpace() || s.contains(QStringLiteral("  "));
}

CellAddress Utility::ParseCoordinate(const QString& coordinate)
{
    CellAddress result {};

    if (coordinate.isEmpty())
        return result;

    const int n = coordinate.size();
    int i = 0;

    // -------------------------
    // 1. parse column letters
    // -------------------------
    int col = 0;
    bool has_column = false;

    while (i < n) {
        const QChar ch = coordinate[i];

        if (ch == u'$') {
            ++i;
            continue;
        }

        if (ch.isLetter()) {
            has_column = true;
            col = col * 26 + (ch.toUpper().unicode() - u'A' + 1);
            ++i;
        } else {
            break;
        }
    }

    if (!has_column)
        return result;

    // -------------------------
    // 2. parse row digits
    // -------------------------
    int row = 0;
    bool has_row = false;

    while (i < n) {
        const QChar ch = coordinate[i];

        if (ch == u'$') {
            ++i;
            continue;
        }

        if (ch.isDigit()) {
            has_row = true;
            row = row * 10 + (ch.unicode() - u'0');
            ++i;
        } else {
            return result; // invalid char
        }
    }

    if (!has_row)
        return result;

    // -------------------------
    // 3. validation (Excel bounds)
    // -------------------------
    if (row <= 0 || col <= 0)
        return result;

    if (row > kMaxExcelRow || col > kMaxExcelColumn)
        return result;

    result.row = row;
    result.column = col;
    result.valid = true;

    return result;
}

QString Utility::ComposeCoordinate(int row, int column, bool row_abs, bool col_abs)
{
    if (row <= 0 || row > kMaxExcelRow || column <= 0 || column > kMaxExcelColumn) {
        qWarning() << "Invalid Excel coordinate:" << row << column;
        return {};
    }

    QString result {};
    result.reserve(16);

    if (col_abs)
        result += QLatin1Char('$');

    result += ComposeColumn(column);

    if (row_abs)
        result += QLatin1Char('$');

    result += QString::number(row);

    return result;
}

QString Utility::ComposeColumn(int column)
{
    if (column <= 0 || column > kMaxExcelColumn) {
        qWarning() << "Invalid Excel column:" << column;
        return {};
    }

    QString col_str {};
    col_str.reserve(4);

    while (column > 0) {
        const int remainder { (column - 1) % 26 };
        col_str.prepend(QChar(u'A' + remainder));
        column = (column - 1) / 26;
    }

    return col_str;
}

YXLSX_END_NAMESPACE
