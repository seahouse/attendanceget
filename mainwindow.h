#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

#include <QJsonArray>
#include <QDateTime>
#include <QMap>

class QNetworkAccessManager;
class QNetworkReply;
class QAxObject;

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
        OTListschedule,                 // 考勤排班详情
        OTListscheduling,               // 考勤排班详情
        OTGetsimplegroups,                 // 考勤组列表详情
        OTGetsimplegroupsing,               // 考勤组列表详情
        OTGetLeaveData,                 // 获取请假数据_token
        OTGetLeaveDataing,              // 获取请假数据
        OTError,                        // 出错
    };

    struct SUserAttendance
    {
        QString _username;
        double _onDuty;                 // 上班打卡天数
        int _lateMinutes;               // 迟到时长（分钟数）
        int _earlyMinutes;              // 早退时长（分钟数）
        int _normalMinutes;             // 工作时长（分钟数）
        int _expectWorkMinutes;         // 预期工作时长（分钟数）
        int _onDutyFull;                // 满勤天数
        double _leaveDays;              // 请假天数
    };

    struct SAttendanceClass
    {
        SAttendanceClass() {}
        SAttendanceClass(double classid, QString classname, int worktimeMinutes) :
            _classid(classid), _classname(classname), _worktimeMinutes(worktimeMinutes) {}
        double _classid;
        QString _classname;
        int _worktimeMinutes;
    };

    struct SAttendanceGroup
    {
        SAttendanceGroup() {}
        SAttendanceGroup(double groupid, QString groupname, QStringList workdayList) :
            _groupid(groupid), _groupname(groupname), _workdayList(workdayList) {}

        double _groupid;
        QString _groupname;
        QStringList _workdayList;
    };

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private slots:
    void sGetAttendance();
    void sListschedule();
    void sGetsimplegroups();
    void sNetworkFinished(QNetworkReply *reply);
    void sTimeout();


private:
    void getToken(OptType ot);

    void listschedule();
    void getsimplegroups();
    void getAttendance2();

    void getDepartment();
    void getUserList();
    void getUserList(int department_id);
    void getAttendance3();

    void handlerExcel();

private slots:
    void sOpenFile();

private slots:
    void sGetLeaveData();       // 获取请假数据

private:
    void getLeaveData(int nextCursor = 0);

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
    QDateTime _dateTimeTo;

    QMap<QString, SUserAttendance> _attendanceDataMap;

    QMap<QString, SAttendanceGroup> _attendanceGroupMap;
    QMap<QString, SAttendanceClass> _attendanceClassMap;
    QMap<QString, double> _leaveDayMap;

    QAxObject *_excel;
};

#endif // MAINWINDOW_H
