#include "mainwindow.h"
#include "log.h"
#include <QApplication>


int main(int argc, char *argv[])
{
    setLogParams("log", "excel_verify", LogLevel::LOG_LEVEL_DEBUG);
    qInstallMessageHandler(customMessageHandler);

    qDebug("ExcelVerify program start ...");

    QApplication a(argc, argv);
    MainWindow w;
    w.show();
    return a.exec();
}
