#ifndef COMMON
#define COMMON

#include <QDir>
#include <QTextStream>

void myMessageOutput(QtMsgType type, const QMessageLogContext &context, const QString &msg)
{
    QDir dir(qApp->applicationDirPath());
    if (!dir.exists("logs"))
        dir.mkdir("logs");
    QFile file(dir.absolutePath() + "/logs/" + QDate::currentDate().toString("yyyy-MM-dd") + ".txt");
    if (file.open(QIODevice::Append | QIODevice::Text))
    {
        QTextStream out(&file);
        QByteArray localMsg = msg.toLocal8Bit();
        switch (type) {
        case QtDebugMsg:
//            out << QString("Debug: %1 (%2:%3, %4)\n")
//                   .arg(localMsg.constData())
//                   .arg(context.file)
//                   .arg(QString::number(context.line))
//                   .arg(context.function);
            fprintf(stdout, "%s\n", localMsg.constData());
            fflush(stdout);
            break;
        case QtWarningMsg:      // use for information
            out << QDateTime::currentDateTime().toString("hh::mm:ss") << QString("%1\n")
                   .arg(msg)
//                   .arg(context.file)
//                   .arg(QString::number(context.line))
//                   .arg(context.function)
                   ;
//            fprintf(stderr, "Fatal: %s (%s:%u, %s)\n", localMsg.constData(), context.file, context.line, context.function);
            break;
        case QtFatalMsg:
            out << QString("Fatal: %1 (%2:%3, %4)\n")
                   .arg(localMsg.constData())
                   .arg(context.file)
                   .arg(QString::number(context.line))
                   .arg(context.function);
            fprintf(stderr, "Fatal: %s (%s:%u, %s)\n", localMsg.constData(), context.file, context.line, context.function);
            break;
        case QtSystemMsg:
            out << QString("Critical: %1 (%2:%3, %4)\n")
                   .arg(localMsg.constData())
                   .arg(context.file)
                   .arg(QString::number(context.line))
                   .arg(context.function);
            fprintf(stderr, "Critical: %s (%s:%u, %s)\n", localMsg.constData(), context.file, context.line, context.function);
            break;
        default:
            break;
        }

        file.close();
    }

}

#endif // COMMON

