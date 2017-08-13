#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

class QNetworkAccessManager;
class QNetworkReply;

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    enum OptType
    {
        OTNone,
        OTGetAccessToken,               // 获取AccessToken
        OTGetAttendance,                // 获取考勤信息
        OTOrderDownloadEnd,             // 订单下载结束
        OTError,                        // 出错
    };

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void sGetAttendance();
    void sNetworkFinished(QNetworkReply *reply);
    void sTimeout();

private:
    void getAttendance();
    void getAttendance2();

private:
    Ui::MainWindow *ui;

    QNetworkAccessManager *_manager;
    QString _token;
    OptType _optType;
    QTimer *_timer;
};

#endif // MAINWINDOW_H
