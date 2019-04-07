/********************************************************************************
** Form generated from reading UI file 'wordtex.ui'
**
** Created by: Qt User Interface Compiler version 5.12.2
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_WORDTEX_H
#define UI_WORDTEX_H

#include <QtCore/QVariant>
#include <QtWidgets/QApplication>
#include <QtWidgets/QMainWindow>
#include <QtWidgets/QMenuBar>
#include <QtWidgets/QStatusBar>
#include <QtWidgets/QToolBar>
#include <QtWidgets/QWidget>

QT_BEGIN_NAMESPACE

class Ui_WordTex
{
public:
    QMenuBar *menuBar;
    QToolBar *mainToolBar;
    QWidget *centralWidget;
    QStatusBar *statusBar;

    void setupUi(QMainWindow *WordTex)
    {
        if (WordTex->objectName().isEmpty())
            WordTex->setObjectName(QString::fromUtf8("WordTex"));
        WordTex->resize(400, 300);
        menuBar = new QMenuBar(WordTex);
        menuBar->setObjectName(QString::fromUtf8("menuBar"));
        WordTex->setMenuBar(menuBar);
        mainToolBar = new QToolBar(WordTex);
        mainToolBar->setObjectName(QString::fromUtf8("mainToolBar"));
        WordTex->addToolBar(mainToolBar);
        centralWidget = new QWidget(WordTex);
        centralWidget->setObjectName(QString::fromUtf8("centralWidget"));
        WordTex->setCentralWidget(centralWidget);
        statusBar = new QStatusBar(WordTex);
        statusBar->setObjectName(QString::fromUtf8("statusBar"));
        WordTex->setStatusBar(statusBar);

        retranslateUi(WordTex);

        QMetaObject::connectSlotsByName(WordTex);
    } // setupUi

    void retranslateUi(QMainWindow *WordTex)
    {
        WordTex->setWindowTitle(QApplication::translate("WordTex", "WordTex", nullptr));
    } // retranslateUi

};

namespace Ui {
    class WordTex: public Ui_WordTex {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_WORDTEX_H
