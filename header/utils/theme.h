// xlsxtheme_p.h

#ifndef YXLSX_THEME_H
#define YXLSX_THEME_H

#include <QIODevice>
#include <QString>

#include "abstractooxmlfile.h"

YXLSX_BEGIN_NAMESPACE

class Theme final : public AbstractOOXmlFile {
public:
    explicit Theme(OperationMode mode);

    void ComposeXml(QIODevice* device) const override;
    bool ParseXml(QIODevice* device) override;
};

YXLSX_END_NAMESPACE

#endif // YXLSX_THEME_H
