#include "mainwindow.h"
#include "ui_mainwindow.h"

#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QTimer>
#include <QJsonObject>
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
    qDebug() << data;
//    ui->textEdit->setText(data);

    QJsonObject json(QJsonDocument::fromJson(data).object());
    int errcode = json.value("errcode").toInt(-1);

    switch (_optType) {
    case OTGetAccessToken:
        if (0 == errcode)
        {
            _token = json.value("access_token").toString();
            qDebug() << _token;

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
    default:
        break;
    }
}

void MainWindow::sTimeout()
{
    _timer->stop();

    switch (_optType) {
    case OTGetAttendance:
        getAttendance();
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

void MainWindow::getAttendance2()
{
    QString url = "https://oapi.dingtalk.com/attendance/listRecord";
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
