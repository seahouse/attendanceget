#include "mainwindow.h"
#include <QApplication>

#include "common.h"

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    qInstallMessageHandler(myMessageOutput);

    if (QDate::currentDate() > QDate(2017, 12, 15))
        return -1;

    MainWindow w;
    w.show();

    return a.exec();
}
