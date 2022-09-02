/*****************************************************************************/
/**
* \file       log.h
* \author     Fanghua Guo
* \date       2022/06/22
* \version
* \brief      ������
* \note       Copyright (c) 2020-2030
* \remarks    �޸���־
******************************************************************************/

#ifndef LOG_H
#define LOG_H
#include <QString>

enum LogLevel
{
	LOG_LEVEL_DEBUG,
	LOG_LEVEL_INFO,
	LOG_LEVEL_WARNING,
	LOG_LEVEL_ERROR,
	LOG_LEVEL_FATAL
};

// �Զ�����־��Ϣ���
void customMessageHandler(QtMsgType type, const QMessageLogContext &context, const QString &msg);

// ����log�ļ���������
void setLogParams(const QString& filePath, const QString& fileName, LogLevel level);

#endif // LOG_H


