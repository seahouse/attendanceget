#include "mainwindow.h"
#include "ui_mainwindow.h"

#include "windows.h"

#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QTimer>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonDocument>
#include <QAxObject>
#include <QFile>
#include <QThread>
#include <QFileDialog>

#include <QDebug>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    _manager = new QNetworkAccessManager;
    connect(_manager, SIGNAL(finished(QNetworkReply*)), this, SLOT(sNetworkFinished(QNetworkReply*)));

    _timer = new QTimer;
    connect(_timer, SIGNAL(timeout()), this, SLOT(sTimeout()));

//    ::CoInitialize(NULL);
    _excel = new QAxObject("Excel.Application");

    connect(ui->pbnGetAttendance, SIGNAL(clicked(bool)), this, SLOT(sGetAttendance()));


    connect(ui->pbnlistschedule, SIGNAL(clicked(bool)), this, SLOT(sListschedule()));
    connect(ui->pbnGetsimplegroups, SIGNAL(clicked(bool)), this, SLOT(sGetsimplegroups()));
    connect(ui->pbnGetLeaveData, SIGNAL(clicked(bool)), this, SLOT(sGetLeaveData()));
    connect(ui->pbnGetAssessData, SIGNAL(clicked(bool)), this, SLOT(sGetAssessData()));
    connect(ui->pbnGetLoanData, SIGNAL(clicked(bool)), this, SLOT(sGetLoanData()));

    connect(ui->pbnOpenExcel, SIGNAL(clicked(bool)), this, SLOT(sOpenFile()));

    ui->dateEdit->setDate(QDate::currentDate());

#ifdef QT_NO_DEBUG
    ui->pbnlistschedule->hide();
    ui->pbnGetsimplegroups->hide();
    ui->pbnGetLeaveData->hide();
    ui->pbnGetAssessData->hide();
    ui->pbnGetLoanData->hide();
#endif
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::getToken(OptType ot)
{
    _optType = ot;
    QNetworkRequest request;

    // 海亚信息技术
//    QString corpid="ding6ed55e00b5328f39";
//    QString corpsecret="gdQvzBl7IW5f3YUSMIkfEIsivOVn8lcXUL_i1BIJvbP4kPJh8SU8B8JuNe8U9JIo";

    // 深圳市彩真包装制品有限公司
    QString corpid="dingf6254688806d21fd35c2f4657eb6378f";
    QString corpsecret="Ljazw5039W0XhF7b7NmRqnll5i_ZD2RHuU054x8w_XWDrO8gcH9qctRESW-LZhyL";

    request.setUrl(QUrl("https://oapi.dingtalk.com/gettoken?corpid=" + corpid + "&corpsecret=" + corpsecret));
    _manager->get(request);
}

void MainWindow::sGetAttendance()
{
    /// 数据获取顺序
    /// 考勤组列表详情
    ///
    ui->teOutput->append("准备统计考勤信息....\n");

    getToken(OTGetAttendance);

//    _optType = OTGetAccessToken;
//    QNetworkRequest request;

//    // 海亚信息技术
////    QString corpid="ding6ed55e00b5328f39";
////    QString corpsecret="gdQvzBl7IW5f3YUSMIkfEIsivOVn8lcXUL_i1BIJvbP4kPJh8SU8B8JuNe8U9JIo";

//    // 深圳市彩真包装制品有限公司
//    QString corpid="dingf6254688806d21fd35c2f4657eb6378f";
//    QString corpsecret="Ljazw5039W0XhF7b7NmRqnll5i_ZD2RHuU054x8w_XWDrO8gcH9qctRESW-LZhyL";

//    request.setUrl(QUrl("https://oapi.dingtalk.com/gettoken?corpid=" + corpid + "&corpsecret=" + corpsecret));
//    _manager->get(request);
}

void MainWindow::sNetworkFinished(QNetworkReply *reply)
{
    QByteArray data = reply->readAll();
//    qDebug() << data;
//    ui->textEdit->setText(data);

    QJsonObject json(QJsonDocument::fromJson(data).object());
    int errcode = json.value("errcode").toInt(-1);

    switch (_optType) {
    case OTGetAccessToken:
        if (0 == errcode)
        {
            _token = json.value("access_token").toString();
//            qDebug() << _token;

            _optType = OTGetAttendance;
            _timer->start(1000);
        }
        else
        {
            _optType = OTError;
            QString errmsg = json.value("errmsg").toString();
//            ui->textEdit->setText(errmsg);
            _timer->start(1000);
        }
        break;
    case OTGetAttendance:
        if (0 == errcode)
        {
            _token = json.value("access_token").toString();

            _optType = OTGetsimplegroupsing;
            _timer->start(1000);
        }
        else
        {
            _optType = OTError;
            QString errmsg = json.value("errmsg").toString();
            ui->teOutput->setText(errmsg);
            _timer->start(1000);
        }
        break;
    case OTGetAttendancing:
        if (0 == errcode)
        {
            _optType = OTGetDepartment;
            _departmentJsonArray = json.value("department").toArray();
            _currentIndex = 0;
            _userIdList.clear();

            foreach (QJsonValue jv, _departmentJsonArray) {
                qDebug() << jv.toObject().value("name").toString();
            }
//            QJsonArray::Iterator it = _departmentJsonArray.begin();
            getUserList();
        }
        break;
    case OTGetDepartment:
        if (0 == errcode)
        {
            QJsonArray userArray = json.value("userlist").toArray();
            foreach (QJsonValue jv, userArray) {
                QString userid = jv.toObject().value("userid").toString();
                qDebug() << jv.toObject().value("userid").toString() << jv.toObject().value("name").toString();
                if (!_userIdList.contains(jv.toObject().value("userid").toString()))
                {
                    _userIdList.append(jv.toObject().value("userid").toString());

//                    _attendanceDataMap[jv.toObject().value("userid").toString()] = SUserAttendance();
                    _attendanceDataMap[jv.toObject().value("userid").toString()]._username = jv.toObject().value("name").toString();
                    _attendanceDataMap[userid]._leaveDays = _leaveDayMap[userid];
                }
            }

            _currentIndex++;
            getUserList();
        }
        else
        {
            QString errmsg = json.value("errmsg").toString();
            qDebug() << errmsg;
        }
        break;
    case OTGetAttendance3:
        if (0 == errcode)
        {
            QJsonArray attendanceDataArray = json.value("recordresult").toArray();
            QList<QDate> dList;
            QList<QDate> dList2;        // 用于统计工作时长
            bool bOnDuty = false;
            bool bOffDuty = false;

//            QList<QDate> dList2;
//            bool bOnDuty2 = false;
//            bool bOffDuty2 = false;
//            bool bLate2 = false;
//            bool bEarly2 = false;

            foreach (QJsonValue jv, attendanceDataArray) {
//                qDebug() << jv.toObject().value("recordId") << jv.toObject().value("workDate");
//                qDebug() << jv.toObject().value("recordId").toDouble() << jv.toObject().value("workDate") << QString::number(jv.toObject().value("workDate").toDouble(), 'f', 0);
//                qDebug() << QDateTime::fromTime_t(jv.toObject().value("workDate").toDouble() / 1000)  << QDateTime::fromTime_t(jv.toObject().value("workDate").toDouble() / 1000);
//                qDebug() << QString::number(jv.toObject().value("recordId").toDouble(), 'f', 0) << QString::number(jv.toObject().value("workDate").toDouble(), 'f', 0) << jv.toObject().value("userId").toString()
//                         << jv.toObject().value("checkType").toString() << jv.toObject().value("timeResult").toString()
//                         << jv.toObject().value("locationResult").toString() << QString::number(jv.toObject().value("baseCheckTime").toDouble(), 'f', 0)
//                         << QString::number(jv.toObject().value("userCheckTime").toDouble(), 'f', 0);
                qDebug() << QString::number(jv.toObject().value("recordId").toDouble(), 'f', 0) << QDateTime::fromTime_t(jv.toObject().value("workDate").toDouble() / 1000).toString("yyyy-MM-dd hh:mm:ss") << jv.toObject().value("userId").toString()
                         << jv.toObject().value("checkType").toString() << jv.toObject().value("timeResult").toString()
                         << jv.toObject().value("locationResult").toString() << QDateTime::fromTime_t(jv.toObject().value("baseCheckTime").toDouble() / 1000).toString("yyyy-MM-dd hh:mm:ss")
                         << QDateTime::fromTime_t(jv.toObject().value("userCheckTime").toDouble() / 1000).toString("yyyy-MM-dd hh:mm:ss")
//                         << QString::number(jv.toObject().value("groupId").toDouble(), 'f', 0) << QString::number(jv.toObject().value("planId").toDouble(), 'f', 0)
                            ;
                if (_attendanceDataMap.contains(jv.toObject().value("userId").toString()))
                {
                    double recordId = jv.toObject().value("recordId").toDouble();
                    QDateTime workDate = QDateTime::fromTime_t(jv.toObject().value("workDate").toDouble() / 1000);
                    QString userId = jv.toObject().value("userId").toString();
                    QString checkType = jv.toObject().value("checkType").toString();
                    QString timeResult = jv.toObject().value("timeResult").toString();
                    QDateTime baseCheckTime = QDateTime::fromTime_t(jv.toObject().value("baseCheckTime").toDouble() / 1000);
                    QDateTime userCheckTime = QDateTime::fromTime_t(jv.toObject().value("userCheckTime").toDouble() / 1000);
                    int approveId = jv.toObject().value("approveId").toInt();
                    if (userId == "091241042840100683")
                    {
                        int i1 = 0;
                    }
                    if (recordId > 0.0 && !dList.contains(workDate.date()))
                    {
                        if (checkType == "OnDuty")
                            bOnDuty = true;
                        if (checkType == "OffDuty")
                            bOffDuty = true;
                        if (bOnDuty && bOffDuty)
                        {
                            _attendanceDataMap[jv.toObject().value("userId").toString()]._onDuty = _attendanceDataMap[jv.toObject().value("userId").toString()]._onDuty + 1;

                            dList.append(QDateTime::fromTime_t(jv.toObject().value("workDate").toDouble() / 1000).date());
                            bOnDuty = false;
                            bOffDuty = false;
                        }
                    }

                    // 统计迟到时长
                    if (recordId > 0.0 && checkType == "OnDuty" && timeResult == "Late")
                    {
                        int lateMinutes = baseCheckTime.secsTo(userCheckTime) / 60;
                        _attendanceDataMap[userId]._lateMinutes = _attendanceDataMap[userId]._lateMinutes + lateMinutes;
                    }

                    // 统计早退时长
                    if (recordId > 0.0 && checkType == "OffDuty" && timeResult == "Early")
                    {
                        int earlyMinutes = userCheckTime.secsTo(baseCheckTime) / 60;
                        if (userCheckTime.secsTo(baseCheckTime) % 60 > 0)
                            earlyMinutes++;
                        _attendanceDataMap[jv.toObject().value("userId").toString()]._earlyMinutes = _attendanceDataMap[jv.toObject().value("userId").toString()]._earlyMinutes + earlyMinutes;
                    }

                    // 统计工作时长
//                    if (recordId > 0.0)
                    {
                        QList<SListschedule> &listscheduleList = _userListscheduleMap[userId]._listscheduleList;

                        // 缺勤，添加到 迟到时长
                        for (int i = 0; i < listscheduleList.size(); i++)
                        {
                            const SListschedule &lv = listscheduleList.at(i);
                            if (lv._checktype == checkType && lv._planchecktime == baseCheckTime)
                            {
                                lv._timeResult = timeResult;
                                lv._userCheckTime = userCheckTime;
                                lv._approveIdAct = approveId;
//                                if (qAbs(recordId) < 0.000001)
//                                {
//                                    if (lv._approveid > 0)
//                                        ;
//                                    else
//                                        ;

//                                }

                            }
                        }
//                        foreach (SListschedule lv, listscheduleList) {
//                            if (lv._checktype == checkType && lv._planchecktime == baseCheckTime)
//                            {
//                                lv._timeResult = timeResult;
//                                lv._userCheckTime = userCheckTime;
//                                int ii = 0;
////                                if (qAbs(recordId) < 0.000001)
////                                {
////                                    if (lv._approveid > 0)
////                                        ;
////                                    else
////                                        ;

////                                }

//                            }
//                        }

                        if (!dList2.contains(workDate.date()))
                        {
                            int weekday = workDate.date().addDays(1).dayOfWeek() - 1;
                            QString groupId = QString::number(jv.toObject().value("groupId").toDouble(), 'f', 0);
                            if (_attendanceGroupMap.contains(groupId))
                            {
                                QString classId = _attendanceGroupMap[groupId]._workdayList.at(weekday);
                                if (!classId.isEmpty() && _attendanceClassMap.contains(classId))
                                {
                                    int worktimeMinutes = _attendanceClassMap[classId]._worktimeMinutes;
//                                    _attendanceDataMap[jv.toObject().value("userId").toString()]._normalMinutes += worktimeMinutes;

                                    // 预期工作时长
//                                    _attendanceDataMap[userId]._expectWorkMinutes += worktimeMinutes;

                                    // 请假的时间
                                    if (_leaveDayMap2.contains(qMakePair(userId, workDate.date())))
                                    {
                                        // 请假时间：请假小时数除以预期工作时长
                                        double leaveDays = _leaveDayMap2[qMakePair(userId, workDate.date())] / (worktimeMinutes / 60.0);
                                        _attendanceDataMap[userId]._leaveDays += leaveDays;
                                    }
                                }
                            }
                            dList2.append(workDate.date());
                        }

                        if (checkType == "OnDuty")
                        {
                            int onworkMinutes = userCheckTime.secsTo(baseCheckTime) / 60;
//                            if (userCheckTime.secsTo(baseCheckTime) % 60 > 0)
//                                onworkMinutes++;
//                            _attendanceDataMap[jv.toObject().value("userId").toString()]._normalMinutes += onworkMinutes;
                        }
                        if (checkType == "OffDuty")
                        {
                            int onworkMinutes = baseCheckTime.secsTo(userCheckTime) / 60;
//                            _attendanceDataMap[jv.toObject().value("userId").toString()]._normalMinutes += onworkMinutes;
                        }

                        qDebug() << _attendanceDataMap[jv.toObject().value("userId").toString()]._normalMinutes;
                    }

                    // 统计工作时长
//                    if (recordId > 0.0)
//                    {
//                        if (!dList2.contains(workDate.date()))
//                        {
//                            int weekday = workDate.date().addDays(1).dayOfWeek() - 1;
//                            QString groupId = QString::number(jv.toObject().value("groupId").toDouble(), 'f', 0);
//                            if (_attendanceGroupMap.contains(groupId))
//                            {
//                                QString classId = _attendanceGroupMap[groupId]._workdayList.at(weekday);
//                                if (!classId.isEmpty() && _attendanceClassMap.contains(classId))
//                                {
//                                    int worktimeMinutes = _attendanceClassMap[classId]._worktimeMinutes;
//                                    _attendanceDataMap[jv.toObject().value("userId").toString()]._normalMinutes += worktimeMinutes;

//                                    // 预期工作时长
//                                    _attendanceDataMap[userId]._expectWorkMinutes += worktimeMinutes;

//                                    // 请假的时间
//                                    if (_leaveDayMap2.contains(qMakePair(userId, workDate.date())))
//                                    {
//                                        // 请假时间：请假小时数除以预期工作时长
//                                        double leaveDays = _leaveDayMap2[qMakePair(userId, workDate.date())] / (worktimeMinutes / 60.0);
//                                        _attendanceDataMap[userId]._leaveDays += leaveDays;
//                                    }
//                                }
//                            }
//                            dList2.append(workDate.date());
//                        }

//                        if (checkType == "OnDuty")
//                        {
//                            int onworkMinutes = userCheckTime.secsTo(baseCheckTime) / 60;
////                            if (userCheckTime.secsTo(baseCheckTime) % 60 > 0)
////                                onworkMinutes++;
//                            _attendanceDataMap[jv.toObject().value("userId").toString()]._normalMinutes += onworkMinutes;
//                        }
//                        if (checkType == "OffDuty")
//                        {
//                            int onworkMinutes = baseCheckTime.secsTo(userCheckTime) / 60;
//                            _attendanceDataMap[jv.toObject().value("userId").toString()]._normalMinutes += onworkMinutes;
//                        }

//                        qDebug() << _attendanceDataMap[jv.toObject().value("userId").toString()]._normalMinutes;
//                    }




                    // 统计满勤天数
//                    if (recordId > 0.0 && !dList2.contains(workDate.date()))
//                    {
//                        if (checkType == "OnDuty")
//                            bOnDuty2 = true;
//                        if (checkType == "OffDuty")
//                            bOffDuty2 = true;
//                    }

                }
            }

            int days = _dateTimeTo.daysTo(QDateTime(ui->dateEdit->date()));
            if (days > 0)
            {
                _dateTimeFrom = _dateTimeFrom.addDays(7);
                if (days >= 7)
                    _dateTimeTo = _dateTimeTo.addDays(7);
                else
                    _dateTimeTo = _dateTimeTo.addDays(days);
                getAttendance3();
            }
            else
            {
                _dateTimeFrom = QDateTime(QDate(ui->dateEdit->date().year(), ui->dateEdit->date().month(), 1));
                // 因为从月初开始计算，不会涉及到加7天后会跨月的情况，所以可以直接加7天
                _dateTimeTo = _dateTimeFrom.addDays(7).addSecs(-1);
                _currentUserIdIndex++;
                getAttendance3();
            }

//            _dateTimeFrom = _dateTimeFrom.addDays(7);
//            if (_dateTimeFrom.msecsTo(QDateTime::currentDateTime()) > 0)
//                getAttendance3();
//            else
//            {
//                _dateTimeFrom = QDateTime(QDate(QDate::currentDate().year(), QDate::currentDate().month(), 1));
//                _currentUserIdIndex++;
//                getAttendance3();
//            }

        }
        else
        {
            QString errmsg = json.value("errmsg").toString();
            qDebug() << errmsg;
        }
        break;
    case OTListschedule:
//        qDebug() << data;
        if (0 == errcode)
        {
            _token = json.value("access_token").toString();
//            qDebug() << _token;

            _optType = OTListscheduling;
            _timer->start(1000);
        }
        else
        {
            _optType = OTError;
            QString errmsg = json.value("errmsg").toString();
//            ui->textEdit->setText(errmsg);
            _timer->start(1000);
        }
        break;
    case OTListscheduling:
        qDebug() << data;
    {
        QJsonObject jo = json.value("dingtalk_smartwork_attends_listschedule_response").toObject().value("result").toObject();
        int ii = jo.value("ding_open_errcode").toInt();
        QString error_msg = jo.value("error_msg").toString();
        if (0 == ii)
        {
            QJsonObject joResult = jo.value("result").toObject();
            bool has_more = joResult.value("has_more").toBool();
            QJsonObject joSchedules = joResult.value("schedules").toObject();
            QJsonArray jaSchedules = joSchedules.value("at_schedule_for_top_vo").toArray();
            foreach (QJsonValue jv, jaSchedules) {
                QJsonObject joSchedule = jv.toObject();
                qDebug() << joSchedule.value("check_type").toString() << joSchedule.value("class_id").toDouble()
                         << joSchedule.value("class_setting_id").toDouble() << joSchedule.value("group_id").toDouble()
                         << joSchedule.value("plan_check_time").toString() << joSchedule.value("userid").toString();
                QString userid = joSchedule.value("userid").toString();
                _userListscheduleMap[userid]._listscheduleList.append(SListschedule(userid,
                                                                                    QString::number(joSchedule.value("group_id").toDouble(), 'f', 0),
                                                                                    QString::number(joSchedule.value("class_id").toDouble(), 'f', 0),
                                                                                    QString::number(joSchedule.value("class_setting_id").toDouble(), 'f', 0),
                                                                                    joSchedule.value("check_type").toString(),
                                                                                    QDateTime::fromString(joSchedule.value("plan_check_time").toString(), "yyyy-MM-dd hh:mm:ss"),
                                                                                    joSchedule.value("approve_id").toInt()));
                if (joSchedule.value("approve_id").toInt() > 0)
                    int i3 = 0;
            }

            if (has_more)
            {
                _listscheduleSize += jaSchedules.size();
                listschedule(_listscheduleSize);
            }
            else    // 第二天或者结束
            {
                if (_listscheduleDatetime.date() < ui->dateEdit->date())
                {
                    _listscheduleDatetime = _listscheduleDatetime.addDays(1);
                    _listscheduleSize = 0;
                    listschedule();
                }
                else
                {
                    _optType = OTGetLeaveDataing;
                    _timer->start(1000);
                }
            }
        }
    }
        break;
    case OTGetsimplegroups:
//        qDebug() << data;
        if (0 == errcode)
        {
            _token = json.value("access_token").toString();
//            qDebug() << _token;

            _optType = OTGetsimplegroupsing;
            _timer->start(1000);
        }
        else
        {
            _optType = OTError;
            QString errmsg = json.value("errmsg").toString();
//            ui->textEdit->setText(errmsg);
            _timer->start(1000);
        }
        break;
    case OTGetsimplegroupsing:
        qDebug() << data;

    {
        QJsonObject jo = json.value("dingtalk_smartwork_attends_getsimplegroups_response").toObject().value("result").toObject();
        int ding_open_errcode = jo.value("ding_open_errcode").toInt();
//        QString error_msg = jo.value("error_msg").toString();
        if (0 == ding_open_errcode)
        {
            QJsonObject joResult = jo.value("result").toObject();
            bool has_more = joResult.value("has_more").toBool();
            QJsonArray jaGroups = joResult.value("groups").toObject().value("at_group_for_top_vo").toArray();
            foreach (QJsonValue jv, jaGroups) {
                QJsonObject joGroup = jv.toObject();
                qDebug() << joGroup.value("group_id").toDouble() << joGroup.value("is_default").toBool()
                         << joGroup.value("group_name").toString() << joGroup.value("default_class_id").toDouble()
                         << joGroup.value("member_count").toInt();
                QJsonArray workdayArray = joGroup.value("work_day_list").toObject().value("string").toArray();
                QStringList workdayList;
                foreach (QJsonValue workdayJson, workdayArray) {
                    workdayList.append(workdayJson.toString());
                }
                _attendanceGroupMap[QString::number(joGroup.value("group_id").toDouble(), 'f', 0)] = SAttendanceGroup(joGroup.value("group_id").toDouble(),
                                                                                                                      joGroup.value("group_name").toString(),
                                                                                                                      workdayList);

                QJsonArray classArray =  joGroup.value("selected_class").toObject().value("at_class_vo").toArray();
                foreach (QJsonValue cl, classArray) {
                    QJsonObject clJson = cl.toObject();
                    QString strClassid = QString::number(clJson.value("class_id").toDouble(), 'f', 0);

                    int minuteTotal = 0;
                    QJsonArray sectionArray = clJson.value("sections").toObject().value("at_section_vo").toArray();
                    foreach (QJsonValue sectionJson, sectionArray) {
                        QJsonArray timeArray = sectionJson.toObject().value("times").toObject().value("at_time_vo").toArray();
                        if (timeArray.size() > 1)
                        {
                            QDateTime dt1 = QDateTime::fromString(timeArray.at(0).toObject().value("check_time").toString(), "yyyy-MM-dd hh:mm:ss");
                            QDateTime dt2 = QDateTime::fromString(timeArray.at(1).toObject().value("check_time").toString(), "yyyy-MM-dd hh:mm:ss");
                            if (dt1.isValid() && dt2.isValid())
                            {
                                minuteTotal += qAbs(dt1.time().secsTo(dt2.time()) / 60);
                            }
                        }
                    }

                    if (!_attendanceClassMap.contains(strClassid))
                        _attendanceClassMap[strClassid] = SAttendanceClass(clJson.value("class_id").toDouble(),
                                                                           clJson.value("class_name").toString(),
//                                                                           clJson.value("setting").toObject().value("work_time_minutes").toInt(),
                                                                           minuteTotal);
                }
            }
        }
        _optType = OTListscheduling;
        _timer->start(1000);
    }
        break;
    case OTGetLeaveData:
        if (0 == errcode)
        {
            _token = json.value("access_token").toString();

            _optType = OTGetLeaveDataing;
            _timer->start(1000);
        }
        else
        {
            _optType = OTError;
            QString errmsg = json.value("errmsg").toString();
            ui->teOutput->append(errmsg);
            _timer->start(1000);
        }
        break;
    case OTGetLeaveDataing:
        qDebug() << data;

    {
        QJsonObject jo = json.value("dingtalk_smartwork_bpms_processinstance_list_response").toObject().value("result").toObject();
        int ding_open_errcode = jo.value("ding_open_errcode").toInt();
        if (0 == ding_open_errcode)
        {
            QJsonObject joResult = jo.value("result").toObject();
            // 游标。给出-1的默认值，当没有该字段时，表示没有更多的数据，用-1表示
            int nextCursor = joResult.value("next_cursor").toInt(-1);
            QJsonArray jaProcess = joResult.value("list").toObject().value("process_instance_top_vo").toArray();
            foreach (QJsonValue jv, jaProcess) {
                QJsonObject joProcess = jv.toObject();

                QString userid = joProcess.value("originator_userid").toString();
                qDebug() << joProcess.value("process_instance_id").toString() << joProcess.value("title").toString()
                         << joProcess.value("create_time").toString() << joProcess.value("finish_time").toString()
                         << joProcess.value("originator_userid").toString() << joProcess.value("originator_dept_id").toString()
                         << joProcess.value("status").toString() << joProcess.value("process_instance_result").toString();
                QJsonArray formValueArray = joProcess.value("form_component_values").toObject().value("form_component_value_vo").toArray();
                if (joProcess.value("status").toString() != "COMPLETED" || joProcess.value("process_instance_result").toString() != "agree")
                    continue;
                foreach (QJsonValue jvFormValue, formValueArray) {
                    QJsonObject joFormValue = jvFormValue.toObject();
                    qDebug() << joFormValue.value("name").toString() << joFormValue.value("value").toString();
                    if (joFormValue.contains("name") && joFormValue.value("name").toString() == "[\"开始时间","结束时间\"]")
                    {
                        QString timeValue = joFormValue.value("value").toString();
                        QStringList timeValueList = timeValue.mid(1, timeValue.size() - 2).split(",", QString::SkipEmptyParts);
                        qDebug() << timeValueList;
                        QJsonDocument jsonDocument = QJsonDocument::fromJson(timeValue.toLatin1());
                        if (jsonDocument.isArray())
                        {
                            QJsonArray timeValueArray = jsonDocument.array();
                            if (timeValueArray.size() > 5)
                            {
                                QString leaveType = timeValueArray.at(3).toString();
                                // 按天请假的，不进行统计，因为在统计考勤时，已经计入
                                if (leaveType == "day")
                                {
//                                    if (_leaveDayMap.contains(userid))
//                                        _leaveDayMap[userid] += timeValueArray.at(2).toInt();
//                                    else
//                                        _leaveDayMap[userid] = timeValueArray.at(2).toInt();
                                }
                                else if (leaveType == "hour")
                                {
                                    QString strDatetime = timeValueArray.at(0).toString();
//                                    qDebug() << timeValueArray.at(0).toString();
                                    QDateTime dt1 = QDateTime::fromString(strDatetime, "yyyy-MM-dd hh:mm");
                                    QDate d1 = dt1.date();
                                    if (d1.isValid())
                                    {
                                        if (_leaveDayMap2.contains(qMakePair(userid, d1)))
                                            _leaveDayMap2[qMakePair(userid, d1)] += timeValueArray.at(2).toDouble();
                                        else
                                            _leaveDayMap2[qMakePair(userid, d1)] = timeValueArray.at(2).toDouble();
                                    }

                                }
                            }
                        }
                    }
                }
            }

            if (nextCursor > 0)
                getLeaveData(nextCursor);
            else
            {
                _optType = OTGetAssessDataing;
                _timer->start(1000);
            }
        }
    }
        break;
    case OTGetAssessData:
        if (0 == errcode)
        {
            _token = json.value("access_token").toString();

            _optType = OTGetAssessDataing;
            _timer->start(1000);
        }
        else
        {
            _optType = OTError;
            QString errmsg = json.value("errmsg").toString();
            ui->teOutput->append(errmsg);
            _timer->start(1000);
        }
        break;
    case OTGetAssessDataing:
        qDebug() << data;

    {
        QJsonObject jo = json.value("dingtalk_smartwork_bpms_processinstance_list_response").toObject().value("result").toObject();
        int ding_open_errcode = jo.value("ding_open_errcode").toInt();
        if (0 == ding_open_errcode)
        {
            QJsonObject joResult = jo.value("result").toObject();
            // 游标。给出-1的默认值，当没有该字段时，表示没有更多的数据，用-1表示
            int nextCursor = joResult.value("next_cursor").toInt(-1);
            QJsonArray jaProcess = joResult.value("list").toObject().value("process_instance_top_vo").toArray();
            foreach (QJsonValue jv, jaProcess) {
                QJsonObject joProcess = jv.toObject();

                QString userid = joProcess.value("originator_userid").toString();
                qDebug() << joProcess.value("process_instance_id").toString() << joProcess.value("title").toString()
                         << joProcess.value("create_time").toString() << joProcess.value("finish_time").toString()
                         << joProcess.value("originator_userid").toString() << joProcess.value("originator_dept_id").toString()
                         << joProcess.value("status").toString() << joProcess.value("process_instance_result").toString();
                QJsonArray formValueArray = joProcess.value("form_component_values").toObject().value("form_component_value_vo").toArray();
                if (joProcess.value("status").toString() != "COMPLETED" || joProcess.value("process_instance_result").toString() != "agree")
                    continue;

                QString name1, name2;
                double amount1 = 0.0, amount2 = 0.0;
                foreach (QJsonValue jvFormValue, formValueArray) {
                    QJsonObject joFormValue = jvFormValue.toObject();
                    qDebug() << joFormValue.value("name").toString() << joFormValue.value("value").toString();
                    if (joFormValue.contains("name"))
                    {
                        QString name = joFormValue.value("name").toString();
                        QString value = joFormValue.value("value").toString();
                        if (name == "奖励人员")
                            name1 = value;
                        if (name == "奖励金额")
                            amount1 = value.toDouble();
                        if (name == "相关人员")
                            name2 = joFormValue.value("value").toString();
                        if (name == "赔偿金额")
                            amount2 = value.toDouble();
                    }
                }

                if (_assessRewardMap.contains(name1))
                    _assessRewardMap[name1] += amount1;
                else
                    _assessRewardMap[name1] = amount1;

                if (_assessPunishMap.contains(name2))
                    _assessPunishMap[name2] += amount2;
                else
                    _assessPunishMap[name2] = amount2;
            }

            if (nextCursor > 0)
                getAssessData(nextCursor);
            else
            {
                _optType = OTGetLoanDataing;
                _timer->start(1000);
            }
        }
    }
        break;

    case OTGetLoanData:
        if (0 == errcode)
        {
            _token = json.value("access_token").toString();

            _optType = OTGetLoanDataing;
            _timer->start(1000);
        }
        else
        {
            _optType = OTError;
            QString errmsg = json.value("errmsg").toString();
            ui->teOutput->append(errmsg);
            _timer->start(1000);
        }
        break;
    case OTGetLoanDataing:
        qDebug() << data;

    {
        QJsonObject jo = json.value("dingtalk_smartwork_bpms_processinstance_list_response").toObject().value("result").toObject();
        int ding_open_errcode = jo.value("ding_open_errcode").toInt();
        if (0 == ding_open_errcode)
        {
            QJsonObject joResult = jo.value("result").toObject();
            // 游标。给出-1的默认值，当没有该字段时，表示没有更多的数据，用-1表示
            int nextCursor = joResult.value("next_cursor").toInt(-1);
            QJsonArray jaProcess = joResult.value("list").toObject().value("process_instance_top_vo").toArray();
            foreach (QJsonValue jv, jaProcess) {
                QJsonObject joProcess = jv.toObject();

                QString userid = joProcess.value("originator_userid").toString();
                qDebug() << joProcess.value("process_instance_id").toString() << joProcess.value("title").toString()
                         << joProcess.value("create_time").toString() << joProcess.value("finish_time").toString()
                         << joProcess.value("originator_userid").toString() << joProcess.value("originator_dept_id").toString()
                         << joProcess.value("status").toString() << joProcess.value("process_instance_result").toString();
                QJsonArray formValueArray = joProcess.value("form_component_values").toObject().value("form_component_value_vo").toArray();
                if (joProcess.value("status").toString() != "COMPLETED" || joProcess.value("process_instance_result").toString() != "agree")
                    continue;

                foreach (QJsonValue jvFormValue, formValueArray) {
                    QJsonObject joFormValue = jvFormValue.toObject();
                    qDebug() << joFormValue.value("name").toString() << joFormValue.value("value").toString();
                    if (joFormValue.contains("name"))
                    {
                        QString name = joFormValue.value("name").toString();
                        QString value = joFormValue.value("value").toString();
                        if (name == "借款金额")
                            _attendanceDataMap[userid]._loanAmount += value.toDouble();
                    }
                }

            }

            if (nextCursor > 0)
                getAssessData(nextCursor);
            else
            {
                _optType = OTGetAttendancing;
                _timer->start(1000);
            }
        }
    }
        break;
    default:
        break;
    }
}

void MainWindow::sTimeout()
{
    _timer->stop();

    switch (_optType) {
    case OTGetAttendancing:
//        getsimplegroups();
//        getAttendance();
        getDepartment();
        break;
//    case OTProductUploadEnd:
//        emit finished(true, tr("上传商品结束。"));
//        break;
    case OTListscheduling:
        _userListscheduleMap.clear();
        _listscheduleSize = 0;
        _listscheduleDatetime = QDateTime(QDate(ui->dateEdit->date().year(), ui->dateEdit->date().month(), 1));
        listschedule();
        break;
    case OTGetsimplegroupsing:
        getsimplegroups();
        break;
    case OTGetLeaveDataing:
        getLeaveData();
        break;
    case OTGetAssessDataing:
        getAssessData();
        break;
    case OTGetLoanDataing:
        getLoanData();
        break;
    default:
        break;
    }
}

void MainWindow::sListschedule()
{
    _userListscheduleMap.clear();
    _listscheduleSize = 0;
    _listscheduleDatetime = QDateTime(QDate(ui->dateEdit->date().year(), ui->dateEdit->date().month(), 1));

    getToken(OTListschedule);
}

void MainWindow::listschedule(int offset)
{
    QString url = "https://eco.taobao.com/router/rest";
    QMap<QString, QString> paramsMap;
    paramsMap["method"] = "dingtalk.smartwork.attends.listschedule";
//    paramsMap["method"] = "dingtalk.smartwork.attends.getsimplegroups";
    paramsMap["session"] = _token;
    paramsMap["timestamp"] = QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");
    paramsMap["format"] = "json";               // 接口返回结果类型:json
    paramsMap["v"] = "2.0";

    paramsMap["work_date"] = _listscheduleDatetime.toString("yyyy-MM-dd hh:mm:ss");
    if (offset > 0)
        paramsMap["offset"] = QString::number(offset);

    QString params;
    QMapIterator<QString, QString> i(paramsMap);
    while (i.hasNext())
    {
        i.next();
        params.append(i.key()).append("=").append(i.value().toLatin1().toPercentEncoding()).append("&");
    }
    url.append("?").append(params);

    QNetworkRequest request;
    request.setUrl(QUrl(url));
    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/x-www-form-urlencoded");
//    req.setHeader(QNetworkRequest::ContentLengthHeader, params.toLatin1().length());

//    request.setUrl(QUrl("https://eco.taobao.com/router/rest" + _token));
//    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");

//    QJsonObject json;
//    json["touser"] = "02656254526756";  // wu: 03016205532, cs: 02656254526756
//    json["toparty"] = "";
//    json["agentid"] = "12064509";       // dp: 11311439, erp: 12064509
//    json["msgtype"] = "text";

//    QJsonObject textJson;
//    textJson["content"] = "这是一条测试消息。";
//    json["text"] = textJson;

//    QJsonDocument jsonDoc(json);
//    QByteArray data = jsonDoc.toJson(QJsonDocument::Compact);
    QByteArray data;

    _manager->post(request, data);
}

void MainWindow::sGetsimplegroups()
{
    getToken(OTGetsimplegroups);
}

void MainWindow::getsimplegroups()
{
    ui->teOutput->append("开始获取考勤组详情....");

    QString url = "https://eco.taobao.com/router/rest";
    QMap<QString, QString> paramsMap;
//    paramsMap["method"] = "dingtalk.smartwork.attends.listschedule";
    paramsMap["method"] = "dingtalk.smartwork.attends.getsimplegroups";
    paramsMap["session"] = _token;
    paramsMap["timestamp"] = QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");
    paramsMap["format"] = "json";               // 接口返回结果类型:json
    paramsMap["v"] = "2.0";

//    paramsMap["work_date"] = QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");

    QString params;
    QMapIterator<QString, QString> i(paramsMap);
    while (i.hasNext())
    {
        i.next();
        params.append(i.key()).append("=").append(i.value().toLatin1().toPercentEncoding()).append("&");
    }
    url.append("?").append(params);

    QNetworkRequest request;
    request.setUrl(QUrl(url));
    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/x-www-form-urlencoded");
//    req.setHeader(QNetworkRequest::ContentLengthHeader, params.toLatin1().length());

//    request.setUrl(QUrl("https://eco.taobao.com/router/rest" + _token));
//    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");

    QByteArray data;

    _manager->post(request, data);
}

void MainWindow::getAttendance2()
{
    QString url = "https://oapi.dingtalk.com/attendance/list";
    QMap<QString, QString> paramsMap;
    paramsMap["access_token"] = _token;

    QString params;
    QMapIterator<QString, QString> i(paramsMap);
    while (i.hasNext())
    {
        i.next();
        params.append(i.key()).append("=").append(i.value().toLatin1().toPercentEncoding()).append("&");
    }
    url.append("?").append(params);

    QNetworkRequest request;
    request.setUrl(QUrl(url));
//    req.setHeader(QNetworkRequest::ContentTypeHeader, "application/x-www-form-urlencoded");
//    req.setHeader(QNetworkRequest::ContentLengthHeader, params.toLatin1().length());

//    request.setUrl(QUrl("https://eco.taobao.com/router/rest" + _token));
//    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");

    QJsonObject json;
    json["touser"] = "02656254526756";  // wu: 03016205532, cs: 02656254526756
    json["toparty"] = "";
    json["agentid"] = "12064509";       // dp: 11311439, erp: 12064509
    json["msgtype"] = "text";

    QJsonObject textJson;
    textJson["content"] = "这是一条测试消息。";
    json["text"] = textJson;

    QJsonDocument jsonDoc(json);
    QByteArray data = jsonDoc.toJson(QJsonDocument::Compact);
//    QByteArray data;

    _manager->post(request, data);
}

void MainWindow::getDepartment()
{
    ui->teOutput->append("开始获取员工信息....\n");
    QString url = "https://oapi.dingtalk.com/department/list";
    QMap<QString, QString> paramsMap;
    paramsMap["access_token"] = _token;

    QString params;
    QMapIterator<QString, QString> i(paramsMap);
    while (i.hasNext())
    {
        i.next();
        params.append(i.key()).append("=").append(i.value().toLatin1().toPercentEncoding()).append("&");
    }
    url.append("?").append(params);

    QNetworkRequest request;
    request.setUrl(QUrl(url));
//    req.setHeader(QNetworkRequest::ContentTypeHeader, "application/x-www-form-urlencoded");
//    req.setHeader(QNetworkRequest::ContentLengthHeader, params.toLatin1().length());

//    request.setUrl(QUrl("https://eco.taobao.com/router/rest" + _token));
//    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");


    _manager->get(request);
}

void MainWindow::getUserList()
{
//    foreach (QJsonValue jv, departmentArray) {
//        qDebug() << jv.toObject().value("name").toString();
//    }

//    if (it != _departmentJsonArray.end())
//    {
//        qDebug() << (*it).toObject().value("name").toString();
//        getUserList((*it).toObject().value("id").toInt());
//    }
//    else
//        qDebug() << "Iterator finished.";

    if (_currentIndex < _departmentJsonArray.size())
        getUserList(_departmentJsonArray.at(_currentIndex).toObject().value("id").toInt());
    else
    {
        _optType = OTGetAttendance3;
        _currentUserIdIndex = 0;
        _dateTimeFrom = QDateTime(QDate(ui->dateEdit->date().year(), ui->dateEdit->date().month(), 1));
        // 因为从月初开始计算，不会涉及到加7天后会跨月的情况，所以可以直接加7天
        _dateTimeTo = _dateTimeFrom.addDays(7).addSecs(-1);
//        _dateTimeTo = QDateTime(QDate(ui->dateEdit->date().year(), ui->dateEdit->date().month(), ui->dateEdit->date().day()), QTime(23, 59, 59));

        ui->teOutput->append("开始获取考勤信息....\n");
        getAttendance3();
    }
}

void MainWindow::getUserList(int department_id)
{
    QString url = "https://oapi.dingtalk.com/user/list";
    QMap<QString, QString> paramsMap;
    paramsMap["access_token"] = _token;
    paramsMap["department_id"] = QString::number(department_id);

    QString params;
    QMapIterator<QString, QString> i(paramsMap);
    while (i.hasNext())
    {
        i.next();
        params.append(i.key()).append("=").append(i.value().toLatin1().toPercentEncoding()).append("&");
    }
    url.append("?").append(params);
    qDebug() << url;

    QNetworkRequest request;
    request.setUrl(QUrl(url));
//    req.setHeader(QNetworkRequest::ContentTypeHeader, "application/x-www-form-urlencoded");
//    req.setHeader(QNetworkRequest::ContentLengthHeader, params.toLatin1().length());

//    request.setUrl(QUrl("https://eco.taobao.com/router/rest" + _token));
//    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");

    _manager->get(request);
}

void MainWindow::getAttendance3()
{
    if (_currentUserIdIndex >= _userIdList.size())
//    if (_currentUserIdIndex >= 5)
    {
        QMapIterator<QString, SUserAttendance> i(_attendanceDataMap);
        while (i.hasNext()) {
            i.next();
            qDebug() << i.key() << ": " << i.value()._username << i.value()._onDuty << i.value()._lateMinutes;
        }
        handlerExcel();
        return;
    }

    QString url = "https://oapi.dingtalk.com/attendance/list";
//    QString url = "https://oapi.dingtalk.com/attendance/listRecord";
    QMap<QString, QString> paramsMap;
    paramsMap["access_token"] = _token;

    QString params;
    QMapIterator<QString, QString> i(paramsMap);
    while (i.hasNext())
    {
        i.next();
        params.append(i.key()).append("=").append(i.value().toLatin1().toPercentEncoding()).append("&");
    }
    url.append("?").append(params);

    QNetworkRequest request;
    request.setUrl(QUrl(url));
//    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/x-www-form-urlencoded");
//    req.setHeader(QNetworkRequest::ContentLengthHeader, params.toLatin1().length());

//    request.setUrl(QUrl("https://eco.taobao.com/router/rest" + _token));
    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");

    QJsonObject json;

    json["userId"] = _userIdList.at(_currentUserIdIndex);
    json["workDateFrom"] = _dateTimeFrom.toString("yyyy-MM-dd hh:mm:ss");
    json["workDateTo"] = _dateTimeTo.toString("yyyy-MM-dd hh:mm:ss");

//    QJsonArray userIds;
//    userIds.append(_userIdList.at(_currentUserIdIndex));
//    json["userIds"] = userIds;
//    json["checkDateFrom"] = _dateTimeFrom.toString("yyyy-MM-dd hh:mm:ss");
//    json["checkDateTo"] = _dateTimeTo.toString("yyyy-MM-dd hh:mm:ss");

    QJsonDocument jsonDoc(json);
    QByteArray data = jsonDoc.toJson(QJsonDocument::Compact);
    qDebug() << data;
//    QByteArray data;

    _manager->post(request, data);
}

void MainWindow::handlerExcel()
{
    // 统计预期工作时长
//    QString checktype;
    QMapIterator<QString, SUserListschedule> iUserListschedule(_userListscheduleMap);
    while (iUserListschedule.hasNext()) {
        iUserListschedule.next();

        const QList<SListschedule> &listscheduleList = iUserListschedule.value()._listscheduleList;

        QDateTime planchecktime;
        QDateTime userCheckTime;
        foreach (SListschedule lv, listscheduleList) {
            if (lv._checktype == "OnDuty")
            {
                planchecktime = lv._planchecktime;
                userCheckTime = lv._userCheckTime;
            }
            if (lv._checktype == "OffDuty")
            {
                // 预期工作时长
                if (planchecktime.isValid())
                {
                    int minutes = planchecktime.secsTo(lv._planchecktime) / 60;
                    _attendanceDataMap[lv._userid]._expectWorkMinutes += minutes;

                    // 缺卡且没有请假，计入迟到时长
//                    if (lv._timeResult == "NotSigned" && lv._approveIdAct <= 0)
//                        _attendanceDataMap[lv._userid]._lateMinutes += minutes;

                    planchecktime = QDateTime(QDate(-1, -1, -1));
                }

                // 实际工作时长 与 缺卡
                if (userCheckTime.isValid())
                {
                    if (lv._timeResult != "NotSigned")
                    {
                        int minutes = userCheckTime.secsTo(lv._userCheckTime) / 60;
                        _attendanceDataMap[lv._userid]._normalMinutes += minutes;
                    }

                    userCheckTime = QDateTime(QDate(-1, -1, -1));
                }

                // 迟到时长

            }
        }
    }


    ui->teOutput->append("开始处理Excel....");
    QString fileName = ui->lineEdit->text().trimmed();
    if (fileName.isEmpty()) return;

    // 将数据转换为按照姓名为关键字
    QMap<QString, SUserAttendance> attendanceDataMapUsername;
    QMapIterator<QString, SUserAttendance> i(_attendanceDataMap);
    while (i.hasNext()) {
        i.next();
        attendanceDataMapUsername[i.value()._username] = i.value();
//        qDebug() << i.key() << ": " << i.value()._username << i.value()._onDuty << i.value()._lateMinutes;
    }

    if (_excel->isNull()) return;
    _excel->dynamicCall("SetVisible", true);
    QAxObject *workbooks = _excel->querySubObject("WorkBooks");
//    connect(workbooks, SIGNAL(exception(int,QString, QString, QString)), this, SLOT(sException(int, QString, QString, QString)));
//    qDebug() << filename;
    QAxObject *workbook = workbooks->querySubObject("Open(QString,QVariant,QVariant)", fileName, 3);
    if (!workbook)
    {
        qDebug() << "Excel file not exists.";
        return;
    }

    QAxObject *ws = workbook->querySubObject("WorkSheets(int)", 1);
    if (!ws)
        qDebug() << "Get worksheet one failed.";
    ws->querySubObject("select");

    QAxObject *usedRange = ws->querySubObject("UsedRange");
    QAxObject *rows = usedRange->querySubObject("Rows");
    QAxObject *columns = usedRange->querySubObject("Columns");

    int iRowStart = usedRange->property("Row").toInt();
    int iColStart = usedRange->property("Column").toInt();
    int iCols = columns->property("Count").toInt();
    int iRows = rows->property("Count").toInt();

    for (int i = iRowStart + 1; i < iRowStart + iRows; i++)
    {
        QAxObject *range = ws->querySubObject("Cells(int, int)", i, iColStart);
        range->querySubObject("select");
        QThread::msleep(500);
        QString username = range->property("Value").toString().trimmed();
        range = ws->querySubObject("Cells(int, int)", i, iColStart + 4);
        QString checksum = range->property("Value").toString();
        range = ws->querySubObject("Cells(int, int)", i, iColStart + 5);
        QString filesize = range->property("Value").toString();

        if (!username.isEmpty())
        {
            qDebug() << username << checksum << filesize;
            if (attendanceDataMapUsername.contains(username))
            {
                // 出勤天数
                QAxObject *cell = ws->querySubObject("Cells(int, int)", i, iColStart + 5);
                cell->querySubObject("SetValue2(QVariant)", attendanceDataMapUsername[username]._onDuty - attendanceDataMapUsername[username]._leaveDays);

                // 迟到时长
                cell = ws->querySubObject("Cells(int, int)", i, iColStart + 8);
                cell->querySubObject("SetValue2(QVariant)", attendanceDataMapUsername[username]._lateMinutes);

                cell = ws->querySubObject("Cells(int, int)", i, iColStart + 9);
                cell->querySubObject("SetValue2(QVariant)", attendanceDataMapUsername[username]._earlyMinutes);

                cell = ws->querySubObject("Cells(int, int)", i, iColStart + 6);
                cell->querySubObject("SetValue2(QVariant)", attendanceDataMapUsername[username]._normalMinutes / 60.0);
//                cell->querySubObject("SetValue2(QVariant)", attendanceDataMapUsername[username]._expectWorkMinutes / 60.0);

                cell = ws->querySubObject("Cells(int, int)", i, iColStart + 13);
                cell->querySubObject("SetValue2(QVariant)", _assessRewardMap[username]);

                cell = ws->querySubObject("Cells(int, int)", i, iColStart + 14);
                cell->querySubObject("SetValue2(QVariant)", _assessPunishMap[username]);

                cell = ws->querySubObject("Cells(int, int)", i, iColStart + 16);
                cell->querySubObject("SetValue2(QVariant)", attendanceDataMapUsername[username]._loanAmount);

                // 月份
                cell = ws->querySubObject("Cells(int, int)", i, iColStart + 1);
                cell->querySubObject("SetValue2(QVariant)", QString::number(ui->dateEdit->date().month()) + "月");
            }

//            filenamesInExcel.append(filename);


//            if (!mapCksum1.contains(filename))
//                error = "There is no " + filename + " in file cksum\n";
//            else
//            {
//                if (mapCksum1.value(filename)._cksum != checksum)
//                    error.append("file " + file.fileName() + ": Cksum check error\n");
//                if (mapCksum1.value(filename)._filesize != filesize)
//                    error.append("file " + file.fileName() + ": filesize check error\n");
//            }
//            if (!error.isEmpty())
//            {
//                qDebug() << error;
//                qWarning() << error;
//                bRtn = false;
//            }
        }
    }
//    workbook->querySubObject("Save()");
    workbook->dynamicCall("Save()");
//    workbook->querySubObject("SaveAs(QVariant)");


//    _excel->setProperty("Visible", true);

//    QAxObject *workbooks = _excel->querySubObject("WorkBooks");
//    QFile file("workbooksDoc.html");
//    if (file.open(QIODevice::WriteOnly | QIODevice::Text))
//    {
//        QTextStream out(&file);
//        out << workbooks->generateDocumentation();
//        file.close();
//    }
//    QVariantList params;
//    QAxObject *workbook = workbooks->querySubObject("Open(QString)", "4345.xlsx");
//    if (!workbook) return;
//    QFile fileWorkbook("workbookDoc.html");
//    if (fileWorkbook.open(QIODevice::WriteOnly | QIODevice::Text))
//    {
//        QTextStream out(&fileWorkbook);
//        out << workbook->generateDocumentation();
//        fileWorkbook.close();
//    }



    workbook->dynamicCall("Close(QVariant)", false);
    workbooks->dynamicCall("Close()");
    _excel->dynamicCall("Quit()");

    ui->teOutput->append("处理结束.");
}

void MainWindow::sOpenFile()
{
    QString fileName = QFileDialog::getOpenFileName(this, tr("Tooling文件"), ".",
                                                    tr("Excel文件(*.xls *.xlsx)"));

    ui->lineEdit->setText(fileName);
}

void MainWindow::sGetLeaveData()
{
    getToken(OTGetLeaveData);
}

void MainWindow::getLeaveData(int nextCursor)
{
    ui->teOutput->append("开始获取请假数据....");

//    if (_currentUserIdIndex >= _userIdList.size())
//    {
//        QMapIterator<QString, SUserAttendance> i(_attendanceDataMap);
//        while (i.hasNext()) {
//            i.next();
//            qDebug() << i.key() << ": " << i.value()._username << i.value()._onDuty << i.value()._lateMinutes;
//        }
//        handlerExcel();
//        return;
//    }

    QString url = "https://eco.taobao.com/router/rest";
    QMap<QString, QString> paramsMap;
    paramsMap["method"] = "dingtalk.smartwork.bpms.processinstance.list";
    paramsMap["session"] = _token;
    paramsMap["timestamp"] = QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");
    paramsMap["format"] = "json";
    paramsMap["v"] = "2.0";

    paramsMap["process_code"] = "PROC-EF6YJDXRN2-KYCJHJ3OM3FW9SVFG93W1-YWTM0L0J-05";
//    uint timeT = QDateTime::currentDateTime().toTime_t();
    uint timeT = QDateTime(QDate(QDate(ui->dateEdit->date().year(), ui->dateEdit->date().month(), 1))).toTime_t();
    paramsMap["start_time"] = QString::number(timeT).append("000");
//    paramsMap["start_time"] = "1502323200000";
    uint timeTEnd = QDateTime(ui->dateEdit->date().addDays(1)).toTime_t();
    paramsMap["end_time"] = QString::number(timeTEnd).append("000");
    if (nextCursor > 0)
        paramsMap["cursor"] = QString::number(nextCursor);

    QString params;
    QMapIterator<QString, QString> i(paramsMap);
    while (i.hasNext())
    {
        i.next();
        params.append(i.key()).append("=").append(i.value().toLatin1().toPercentEncoding()).append("&");
    }
    url.append("?").append(params);

    QNetworkRequest request;
    request.setUrl(QUrl(url));
//    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/x-www-form-urlencoded");
//    req.setHeader(QNetworkRequest::ContentLengthHeader, params.toLatin1().length());

//    request.setUrl(QUrl("https://eco.taobao.com/router/rest" + _token));
    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");

//    QJsonObject json;
//    json["userId"] = _userIdList.at(_currentUserIdIndex);

//    json["workDateFrom"] = _dateTimeFrom.toString("yyyy-MM-dd hh:mm:ss");
//    json["workDateTo"] = _dateTimeTo.toString("yyyy-MM-dd hh:mm:ss");

//    QJsonDocument jsonDoc(json);
//    QByteArray data = jsonDoc.toJson(QJsonDocument::Compact);
//    qDebug() << data;
    QByteArray data;

    _manager->post(request, data);
}

void MainWindow::sGetAssessData()
{
    getToken(OTGetAssessData);
}

void MainWindow::getAssessData(int nextCursor)
{
    ui->teOutput->append("开始获取每天考核表....");

//    if (_currentUserIdIndex >= _userIdList.size())
//    {
//        QMapIterator<QString, SUserAttendance> i(_attendanceDataMap);
//        while (i.hasNext()) {
//            i.next();
//            qDebug() << i.key() << ": " << i.value()._username << i.value()._onDuty << i.value()._lateMinutes;
//        }
//        handlerExcel();
//        return;
//    }

    QString url = "https://eco.taobao.com/router/rest";
    QMap<QString, QString> paramsMap;
    paramsMap["method"] = "dingtalk.smartwork.bpms.processinstance.list";
    paramsMap["session"] = _token;
    paramsMap["timestamp"] = QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");
    paramsMap["format"] = "json";
    paramsMap["v"] = "2.0";

    paramsMap["process_code"] = "PROC-QNALQBPV-0EXMF94DMBIXD2IFYANU3-YO42TL5J-H";
//    uint timeT = QDateTime::currentDateTime().toTime_t();
    uint timeT = QDateTime(QDate(QDate(ui->dateEdit->date().year(), ui->dateEdit->date().month(), 1))).toTime_t();
    paramsMap["start_time"] = QString::number(timeT).append("000");
//    paramsMap["start_time"] = "1502323200000";
    // 结束日期，修改为当前日期
//    uint timeTEnd = QDateTime(ui->dateEdit->date().addDays(1)).toTime_t();
    uint timeTEnd = QDateTime::currentDateTime().toTime_t();
    paramsMap["end_time"] = QString::number(timeTEnd).append("000");
    if (nextCursor > 0)
        paramsMap["cursor"] = QString::number(nextCursor);

    QString params;
    QMapIterator<QString, QString> i(paramsMap);
    while (i.hasNext())
    {
        i.next();
        params.append(i.key()).append("=").append(i.value().toLatin1().toPercentEncoding()).append("&");
    }
    url.append("?").append(params);

    QNetworkRequest request;
    request.setUrl(QUrl(url));
//    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/x-www-form-urlencoded");
//    req.setHeader(QNetworkRequest::ContentLengthHeader, params.toLatin1().length());

//    request.setUrl(QUrl("https://eco.taobao.com/router/rest" + _token));
    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");

//    QJsonObject json;
//    json["userId"] = _userIdList.at(_currentUserIdIndex);

//    json["workDateFrom"] = _dateTimeFrom.toString("yyyy-MM-dd hh:mm:ss");
//    json["workDateTo"] = _dateTimeTo.toString("yyyy-MM-dd hh:mm:ss");

//    QJsonDocument jsonDoc(json);
//    QByteArray data = jsonDoc.toJson(QJsonDocument::Compact);
//    qDebug() << data;
    QByteArray data;

    _manager->post(request, data);
}

void MainWindow::sGetLoanData()
{
    ui->teOutput->append("开始获取借款数据....");
    getToken(OTGetLoanData);
}

void MainWindow::getLoanData(int nextCursor)
{

//    if (_currentUserIdIndex >= _userIdList.size())
//    {
//        QMapIterator<QString, SUserAttendance> i(_attendanceDataMap);
//        while (i.hasNext()) {
//            i.next();
//            qDebug() << i.key() << ": " << i.value()._username << i.value()._onDuty << i.value()._lateMinutes;
//        }
//        handlerExcel();
//        return;
//    }

    QString url = "https://eco.taobao.com/router/rest";
    QMap<QString, QString> paramsMap;
    paramsMap["method"] = "dingtalk.smartwork.bpms.processinstance.list";
    paramsMap["session"] = _token;
    paramsMap["timestamp"] = QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");
    paramsMap["format"] = "json";
    paramsMap["v"] = "2.0";

    paramsMap["process_code"] = "PROC-FF6Y696SO2-7YCJV8COOFAJQM7F51RV1-UA3SCN0J-HB";
//    uint timeT = QDateTime::currentDateTime().toTime_t();
    uint timeT = QDateTime(QDate(QDate(ui->dateEdit->date().year(), ui->dateEdit->date().month(), 1))).toTime_t();
    paramsMap["start_time"] = QString::number(timeT).append("000");
//    paramsMap["start_time"] = "1502323200000";
    uint timeTEnd = QDateTime(ui->dateEdit->date().addDays(1)).toTime_t();
    paramsMap["end_time"] = QString::number(timeTEnd).append("000");
    if (nextCursor > 0)
        paramsMap["cursor"] = QString::number(nextCursor);

    QString params;
    QMapIterator<QString, QString> i(paramsMap);
    while (i.hasNext())
    {
        i.next();
        params.append(i.key()).append("=").append(i.value().toLatin1().toPercentEncoding()).append("&");
    }
    url.append("?").append(params);

    QNetworkRequest request;
    request.setUrl(QUrl(url));
//    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/x-www-form-urlencoded");
//    req.setHeader(QNetworkRequest::ContentLengthHeader, params.toLatin1().length());

//    request.setUrl(QUrl("https://eco.taobao.com/router/rest" + _token));
    request.setHeader(QNetworkRequest::ContentTypeHeader, "application/json");

//    QJsonObject json;
//    json["userId"] = _userIdList.at(_currentUserIdIndex);

//    json["workDateFrom"] = _dateTimeFrom.toString("yyyy-MM-dd hh:mm:ss");
//    json["workDateTo"] = _dateTimeTo.toString("yyyy-MM-dd hh:mm:ss");

//    QJsonDocument jsonDoc(json);
//    QByteArray data = jsonDoc.toJson(QJsonDocument::Compact);
//    qDebug() << data;
    QByteArray data;

    _manager->post(request, data);
}
