#ifndef WORDTEX_H
#define WORDTEX_H

#include <QMainWindow>

namespace Ui {
class WordTex;
}

class WordTex : public QMainWindow
{
    Q_OBJECT

public:
    explicit WordTex(QWidget *parent = nullptr);
    ~WordTex();

private:
    Ui::WordTex *ui;
};

#endif // WORDTEX_H
