#include "wordtex.h"
#include "ui_wordtex.h"

WordTex::WordTex(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::WordTex)
{
    ui->setupUi(this);
}

WordTex::~WordTex()
{
    delete ui;
}
