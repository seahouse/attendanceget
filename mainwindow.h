#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

#include <QJsonArray>
#include <QDateTime>
#include <QMap>

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
        OTGetDepartment,                // 获取部门与成员信息
        OTGetAttendance3,
        OTError,                        // 出错
    };

    struct SUserAttendance
    {
        QString _username;
        int _onDuty;                    // 上班打卡天数
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

    void getDepartment();
    void getUserList();
    void getUserList(int department_id);
    void getAttendance3();

private:
    Ui::MainWindow *ui;

    QNetworkAccessManager *_manager;
    QString _token;
    OptType _optType;
    QTimer *_timer;

    QJsonArray _departmentJsonArray;
    int _currentIndex;

    QStringList _userIdList;
    int _currentUserIdIndex;
    QDateTime _dateTimeFrom;

    QMap<QString, SUserAttendance> _attendanceDataMap;
};

#endif // MAINWINDOW_H
