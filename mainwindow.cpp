#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QTimer>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonDocument>

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

    connect(ui->pbnGetAttendance, SIGNAL(clicked(bool)), this, SLOT(sGetAttendance()));
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::sGetAttendance()
{
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
                         << QDateTime::fromTime_t(jv.toObject().value("userCheckTime").toDouble() / 1000).toString("yyyy-MM-dd hh:mm:ss");
                if (_attendanceDataMap.contains(jv.toObject().value("userId").toString()))
                {
                    double recordId = jv.toObject().value("recordId").toDouble();
                    QDateTime workDate = QDateTime::fromTime_t(jv.toObject().value("workDate").toDouble() / 1000);
                    if (recordId > 0.0 && !dList.contains(workDate.date()))
                    {
                        QString checkType = jv.toObject().value("checkType").toString();
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

                }
            }

            _dateTimeFrom = _dateTimeFrom.addDays(7);
            if (_dateTimeFrom.msecsTo(QDateTime::currentDateTime()) > 0)
                getAttendance3();
            else
            {
                _dateTimeFrom = QDateTime(QDate(QDate::currentDate().year(), QDate::currentDate().month(), 1));
                _currentUserIdIndex++;
                getAttendance3();
            }

        }
        else
        {
            QString errmsg = json.value("errmsg").toString();
            qDebug() << errmsg;
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
    default:
        break;
    }
}

void MainWindow::getAttendance()
{
    QString url = "https://eco.taobao.com/router/rest";
    QMap<QString, QString> paramsMap;
//    paramsMap["method"] = "dingtalk.smartwork.attends.listschedule";
    paramsMap["method"] = "dingtalk.smartwork.attends.getsimplegroups";
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
        _dateTimeFrom = QDateTime(QDate(QDate::currentDate().year(), QDate::currentDate().month(), 1));
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
    {
        QMapIterator<QString, SUserAttendance> i(_attendanceDataMap);
        while (i.hasNext()) {
            i.next();
            qDebug() << i.key() << ": " << i.value()._username << i.value()._onDuty << endl;
        }
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

    QDateTime dateTime = QDateTime::currentDateTime();
//    json["workDateFrom"] = QDateTime(QDate(dateTime.date().year(), dateTime.date().month(), 1)).toString("yyyy-MM-dd hh:mm:ss");
//    json["workDateTo"] = QDateTime::currentDateTime().toString("yyyy-MM-dd hh:mm:ss");
    json["workDateFrom"] = _dateTimeFrom.toString("yyyy-MM-dd hh:mm:ss");
    json["workDateTo"] = _dateTimeFrom.addDays(7).addSecs(-1).toString("yyyy-MM-dd hh:mm:ss");

//    QJsonObject textJson;
//    textJson["content"] = "这是一条测试消息。";
//    json["text"] = textJson;

    QJsonDocument jsonDoc(json);
    QByteArray data = jsonDoc.toJson(QJsonDocument::Compact);
    qDebug() << data;
//    QByteArray data;

    _manager->post(request, data);
}
