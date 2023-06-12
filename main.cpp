#include"form.h"
#include <QApplication>
#include<QSplashScreen>
#include<QSqlDatabase>
#include<QtDebug>
#include<QSqlError>
#include<QMouseEvent>
int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    QString rccpath = "E:\\QT\\CH\\RAMS\\rams\\images.rcc";
    QResource::registerResource(rccpath);

    //引用rcc文件的图片
    QIcon icon(QPixmap::fromImage(QImage(":LOGO")));

    //连接MYSQL数据库
    Eigen::Vector2d v;
    v << 2, 4;
    QSqlDatabase db = QSqlDatabase::addDatabase("QMYSQL");
    QStringList list= QSqlDatabase::drivers();
    qDebug()<<list;
    db.setHostName("localhost");
    db.setUserName("root");
    db.setDatabaseName("mydb");
    db.setPassword("185679");
    db.setPort(3306);
    if(db.open())
    {
        qDebug()<<"open success!!!!";
    }
    else
    {
        qDebug()<<"open failed" << db.lastError().text();
    }
    QPixmap pixmap("metro_2");
    pixmap=pixmap.scaled(800,600,Qt::KeepAspectRatio);
    QSplashScreen splash(pixmap);
    splash.show();
    a.processEvents();
    Form w;
    w.show();
    splash.finish(&w);
    return a.exec();
}
