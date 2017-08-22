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

    connect(ui->pbnOpenExcel, SIGNAL(clicked(bool)), this, SLOT(sOpenFile()));

    ui->dateEdit->setDate(QDate::currentDate());

#ifdef QT_NO_DEBUG
    ui->pbnlistschedule->hide();
    ui->pbnGetsimplegroups->hide();
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
    ui->teOutput->append("开始统计....\n");
    _optType = OTGetAccessToken;
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
                qDebug() << jv.toObject().value("userid").toString() << jv.toObject().value("name").toString();
                if (!_userIdList.contains(jv.toObject().value("userid").toString()))
                {
                    _userIdList.append(jv.toObject().value("userid").toString());

                    _attendanceDataMap[jv.toObject().value("userid").toString()] = SUserAttendance();
                    _attendanceDataMap[jv.toObject().value("userid").toString()]._username = jv.toObject().value("name").toString();
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
                         << QString::number(jv.toObject().value("groupId").toDouble(), 'f', 0) << QString::number(jv.toObject().value("planId").toDouble(), 'f', 0);
                if (_attendanceDataMap.contains(jv.toObject().value("userId").toString()))
                {
                    double recordId = jv.toObject().value("recordId").toDouble();
                    QDateTime workDate = QDateTime::fromTime_t(jv.toObject().value("workDate").toDouble() / 1000);
                    QString checkType = jv.toObject().value("checkType").toString();
                    QString timeResult = jv.toObject().value("timeResult").toString();
                    QDateTime baseCheckTime = QDateTime::fromTime_t(jv.toObject().value("baseCheckTime").toDouble() / 1000);
                    QDateTime userCheckTime = QDateTime::fromTime_t(jv.toObject().value("userCheckTime").toDouble() / 1000);
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
                        _attendanceDataMap[jv.toObject().value("userId").toString()]._lateMinutes = _attendanceDataMap[jv.toObject().value("userId").toString()]._lateMinutes + lateMinutes;
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
                    if (recordId > 0.0 && checkType == "OffDuty" && timeResult == "Early")
                    {
                        int earlyMinutes = userCheckTime.secsTo(baseCheckTime) / 60;
                        if (userCheckTime.secsTo(baseCheckTime) % 60 > 0)
                            earlyMinutes++;
                        _attendanceDataMap[jv.toObject().value("userId").toString()]._earlyMinutes = _attendanceDataMap[jv.toObject().value("userId").toString()]._earlyMinutes + earlyMinutes;
                    }

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

//        QJsonObject jo(QJsonDocument::fromJson(json.value("dingtalk_smartwork_attends_listschedle_response")).object());
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
            }

            int aa = 0;
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
            }

            int aa = 0;
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
    case OTGetAttendance:
//        getAttendance();
        getDepartment();
        break;
//    case OTProductUploadEnd:
//        emit finished(true, tr("上传商品结束。"));
//        break;
    case OTListscheduling:
        listschedule();
        break;
    case OTGetsimplegroupsing:
        getsimplegroups();
        break;
    default:
        break;
    }
}

void MainWindow::sListschedule()
{
    getToken(OTListschedule);
}

void MainWindow::listschedule()
{
    QString url = "https://eco.taobao.com/router/rest";
    QMap<QString, QString> paramsMap;
    paramsMap["method"] = "dingtalk.smartwork.attends.listschedule";
//    paramsMap["method"] = "dingtalk.smartwork.attends.getsimplegroups";
    paramsMap["session"] = _token;
    paramsMap["timestamp"] = QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");
    paramsMap["format"] = "json";               // 接口返回结果类型:json
    paramsMap["v"] = "2.0";

    paramsMap["work_date"] = QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");

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

//    QDateTime dateTime = QDateTime::currentDateTime();
//    json["workDateFrom"] = QDateTime(QDate(dateTime.date().year(), dateTime.date().month(), 1)).toString("yyyy-MM-dd hh:mm:ss");
//    json["workDateTo"] = QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");
    json["workDateFrom"] = _dateTimeFrom.toString("yyyy-MM-dd hh:mm:ss");
    json["workDateTo"] = _dateTimeTo.toString("yyyy-MM-dd hh:mm:ss");
//    json["workDateTo"] = _dateTimeFrom.addDays(7).addSecs(-1).toString("yyyy-MM-dd hh:mm:ss");

//    QJsonObject textJson;
//    textJson["content"] = "这是一条测试消息。";
//    json["text"] = textJson;

    QJsonDocument jsonDoc(json);
    QByteArray data = jsonDoc.toJson(QJsonDocument::Compact);
    qDebug() << data;
//    QByteArray data;

    _manager->post(request, data);
}

void MainWindow::handlerExcel()
{
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
                QAxObject *cell = ws->querySubObject("Cells(int, int)", i, iColStart + 5);
                cell->querySubObject("SetValue2(QVariant)", attendanceDataMapUsername[username]._onDuty);

                cell = ws->querySubObject("Cells(int, int)", i, iColStart + 8);
                cell->querySubObject("SetValue2(QVariant)", attendanceDataMapUsername[username]._lateMinutes);

                cell = ws->querySubObject("Cells(int, int)", i, iColStart + 9);
                cell->querySubObject("SetValue2(QVariant)", attendanceDataMapUsername[username]._earlyMinutes);
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
