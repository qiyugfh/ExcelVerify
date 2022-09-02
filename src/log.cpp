/*****************************************************************************/
/**
* \file       log.cpp
* \author     Fanghua Guo
* \date       2022/06/22
* \version
* \brief
* \note       Copyright (c) 2020-2030
* \remarks    �޸���־
******************************************************************************/

#include "log.h"
#include <QCoreApplication>
#include <QDir>
#include <QDateTime>
#include <QMutex>
#include <QTextStream>
#include <QThread>


static LogLevel s_logLevel = LogLevel::LOG_LEVEL_DEBUG;
static QString  s_logFileName = "";
static QString  s_logFilePath = ".";
static QMutex   s_mutex;

/*****************************************************************************/
/**
* \author     Fanghua Guo
* \date       2022/06/22
* \brief
* \param[in]  info   ��־�ļ��𣬺�����������Ϣ�����
* \param[in]  msg    ��־����
* \return
* \remarks    ������Ϣ
******************************************************************************/
void outputMsg(const QString& info, const QString& msg)
{
    QString out = QString("[%1]%2: %3").arg(QDateTime::currentDateTime().toString("yyyy-MM-dd HH:mm:ss.zzz")).arg(info).arg(msg);

    // ��ӡ���������־�ļ���
    s_mutex.lock();
    QFile logFile(QString("%1/%2.log").arg(s_logFilePath).arg(s_logFileName));
    if (logFile.open(QIODevice::WriteOnly | QIODevice::Append))
    {
        QTextStream stream(&logFile);
        stream << out << endl;
        logFile.flush();
        logFile.close();
    }
    s_mutex.unlock();

#ifdef QT_DEBUG
    printf("%s\n", out.toStdString().c_str());
    fflush(stdout);
#endif

}


/*****************************************************************************/
/**
* \author     Fanghua Guo
* \date       2022/06/22
* \brief
* \param[in]  type    ��־����
* \param[in]  context ��־������
* \param[in]  msg     ��־����
* \return
* \remarks    ������Ϣ
******************************************************************************/
void customMessageHandler(QtMsgType type, const QMessageLogContext &context, const QString &msg)
{
    QString fileName = QString(context.file);
    if ((NULL != fileName)
        && (0 < fileName.length()))
    {
        int pos = fileName.lastIndexOf("\\");
        if (0 < pos)
        {
            fileName = fileName.right(fileName.length() - pos - 1);
        }
    }

    QString info;
    switch (type)
    {
    case QtDebugMsg:
        if(int(LogLevel::LOG_LEVEL_DEBUG) < int(s_logLevel))
        {
            return;
        }
        info = "[debug]";
        break;
    case QtWarningMsg:
        if (int(LogLevel::LOG_LEVEL_WARNING) < int(s_logLevel))
        {
            return;
        }
        info = "[warn ]";
        break;
    case QtCriticalMsg:
        if (int(LogLevel::LOG_LEVEL_ERROR) < int(s_logLevel))
        {
            return;
        }
        info = "[error]";
        break;
    case QtFatalMsg:
        if (int(LogLevel::LOG_LEVEL_FATAL) < int(s_logLevel))
        {
            return;
        }
        info = "[fatal]";
        break;
    case QtInfoMsg:
        if (int(LogLevel::LOG_LEVEL_INFO) < int(s_logLevel))
        {
            return;
        }
        info = "[info ]";
        break;
    default:
        info = "[unkown]";
    }

    QString threadId = QString("0x%1").arg(quintptr(QThread::currentThreadId()), 0, 16);
    info += QString("[debug][%1]<%2:%3:%4>").arg(threadId).arg(fileName).arg(context.function).arg(context.line);
    outputMsg(info, msg);
}

void setLogParams(const QString &filePath, const QString& fileName, LogLevel level)
{
    s_logFilePath = filePath;
    s_logFileName = fileName;
    s_logLevel = level;

    QDir dir;
    if (!dir.exists(filePath))
    {
        dir.mkdir(filePath);
    }
}
