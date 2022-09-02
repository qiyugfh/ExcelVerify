/*****************************************************************************/
/**
* \file       log.h
* \author     Fanghua Guo
* \date       2022/06/22
* \version
* \brief      主界面
* \note       Copyright (c) 2020-2030
* \remarks    修改日志
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

// 自定义日志消息句柄
void customMessageHandler(QtMsgType type, const QMessageLogContext &context, const QString &msg);

// 设置log文件名、级别
void setLogParams(const QString& filePath, const QString& fileName, LogLevel level);

#endif // LOG_H


