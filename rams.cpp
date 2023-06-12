#include "rams.h"
#include "ui_rams.h"
#include<QSqlQueryModel>
#include<Eigen/Dense>
#include<QSqlQuery>
using namespace Eigen;
bool Rams::on_Pushbutton_Flag_1 = false;
bool Rams::on_Pushbutton_Flag_2 = false;
bool Rams::on_Pushbutton_Flag_3 = false;
Rams::Rams(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::Rams)
{
    //去掉边框
    setWindowFlag(Qt::WindowType::FramelessWindowHint);
    //背景透明
    setAttribute(Qt::WA_TranslucentBackground);

    ui->setupUi(this);

    QSqlQuery query;
    query.exec("select * from routput");
    query.last();
    ui->lineEdit->setText(query.value(0).toString());
    ui->lineEdit_2->setText(query.value(1).toString());
    myexcel = new QAxObject("Excel.Application");
    myworks = myexcel->querySubObject("WorkBooks");
    myworks -> dynamicCall("Add");
    workbook = myexcel ->querySubObject("ActiveWorkBook");
    mysheets = workbook->querySubObject("Sheets");

    //将节点输入到Nodes中
    QSqlQueryModel *FtaModel = new QSqlQueryModel(this);
    FtaModel->setQuery("select BottomEvent from hlp");
    ui->Nodes->setModel(FtaModel);

    //将部件输入到parts
    QStringList list;
    list<<"汇流排"<<"接触线"<<"绝缘子"<<"定位装置"<<"分段绝缘器"<<"上网引线";
    ui->parts->addItems(list);

    //将子系统输入到parts_2中
    QStringList strList;
    query.exec("select * from mohu");
    while(query.next())
    {
       if(query.value(0).toString()!=NULL)
      { strList<<query.value(0).toString();}
       else
           break;
    }
    ui->parts_2->addItems(strList);

    //将因素输入Factor中
    QStringList nameList;
    nameList<<"人员因素"<<"设备因素"<<"环境因素"<<"管理因素";
    ui->Factor->addItems(nameList);

    //将具体因素输入factors中
    QSqlQueryModel *FactorModel = new QSqlQueryModel(this);
    FactorModel->setQuery("select Human from human");
    ui->factors->setModel(FactorModel);

    //提示
   // ui->Nodes->lineEdit()->setPlaceholderText("节点名称");

}

Rams::~Rams()
{
    delete ui;
}
//函数计算
double Rams::countRability(QVariantList list, int n)
{
    double a,r=1;

    for (int i=0;i<n;i++) {
        a=list.at(i).toDouble();
        a=1-a;
        r=r*a;
      // qDebug()<<(QString::number(r,10,10));
    }
    QSqlQuery query;
    query.exec(QString("upadate routput set 可靠值='%1'").arg(QString::number(r)));
    query.exec(QString("upadate routput set 故障间隔时间='%2'").arg(QString::number(1/(1-r))));
    return r;
}

double Rams::countProImportance(QString NodeName, QVariantList list, double r)
{
    double P=0;

    if (NodeName=="X1"){
    P=r/(1-list.at(0).toDouble());
    }
    else if (NodeName=="X2") {
    P=r/(1-list.at(1).toDouble());
    }
    else if (NodeName=="X3") {
    P=r/(1-list.at(2).toDouble());
    }
    else if (NodeName=="X4") {
    P=r/(1-list.at(3).toDouble());
    }
    else if (NodeName=="X5") {
    P=r/(1-list.at(4).toDouble());
    }
    else if (NodeName=="X6") {
    P=r/(1-list.at(5).toDouble());
    }
    else if (NodeName=="X7") {
    P=r/(1-list.at(6).toDouble());
    }
    else if (NodeName=="X8") {
    P=r/(1-list.at(7).toDouble());
    }
    else if (NodeName=="X9") {
    P=r/(1-list.at(8).toDouble());
    }
    else if (NodeName=="X10") {
    P=r/(1-list.at(9).toDouble());
    }
    else if (NodeName=="X11") {
    P=r/(1-list.at(10).toDouble());
    }
    else if (NodeName=="X12") {
    P=r/(1-list.at(11).toDouble());
    }
    else if (NodeName=="X13") {
    P=r/(1-list.at(12).toDouble());
    }
    else if (NodeName=="X14") {
    P=r/(1-list.at(13).toDouble());
    }
    else if (NodeName=="X15") {
    P=r/(1-list.at(14).toDouble());
    }
    else if (NodeName=="X16") {
    P=r/(1-list.at(15).toDouble());
    }
    else if (NodeName=="X17") {
    P=r/(1-list.at(16).toDouble());
    }
    else if (NodeName=="X18") {
    P=r/(1-list.at(17).toDouble());
    }
    else if (NodeName=="X19") {
    P=r/(1-list.at(18).toDouble());
    }
    else if (NodeName=="X20") {
    P=r/(1-list.at(19).toDouble());
    }

    return P;
}

float Rams::count_DisRe(QString name)
{
    QSqlQuery query;
    query.exec("select * from mohu");
    QList<float> list,list1,list2,list3;
    float b=0;
    MatrixXf D(4,6),R(6,4),A(1,6),E(1,4),F(1,4);
    F=MatrixXf::Ones(1,4);
    while(query.next())
    {
        list1<<query.value("FScore_1").toFloat();
        list2<<query.value("FScore_2").toFloat();
        list3<<query.value("FScore_3").toFloat();
    }
    for (int i=0;i<=5;i++) {
        list<<(list1.at(i)+list2.at(i)+list3.at(i))/3;
        b=b+list.at(i);
    }
    for (int i=0;i<=5;i++) {
        list[i]=list[i]/b;
        A(i) = list[i];
    }
    for (int i=8;i<=17;i=i+3) {


    int J=0;
    query.first();
    while(J<=5)
    {
        float l,o,p;
        if(J==0)
            query.first();
        l=query.value(i).toFloat();
        o=query.value(i+1).toFloat();
        p=query.value(i+2).toFloat();
        D((i-8)/3,J)=(query.value(i).toFloat()+query.value(i+1).toFloat()+query.value(i+2).toFloat())/30;
        J++;
        query.next();
    }
    }

    R=D.transpose();
    E=A*R;
    F=F-E;
    float Rs=0.98,r;
    QString rD= QString("%1/%2/%3/%4").arg(QString::number(pow(Rs,F(0,0)))).arg(QString::number(pow(Rs,F(0,1)))).arg(QString::number(pow(Rs,F(0,2)))).arg(QString::number(pow(Rs,F(0,3))));
    query.exec(QString("update routput set 可靠度分配='%1'").arg(rD));
    if(name=="接触悬挂")
    {
        r=pow(Rs,F(0,0));
        return r;
    }
    else if(name=="支持定位")
    {
        r=pow(Rs,F(0,1));
        return r;
    }
    else if(name=="绝缘装置")
    {
        r=pow(Rs,F(0,2));
        return r;
    }
    else if(name=="附加导线")
    {
        r=pow(Rs,F(0,3));
        return r;
    }
}

QString Rams::count_Level(QString arg1, QString arg2)
{
    QString leve_1="极高",leve_2="高",leve_3="中",leve_4="低";

    if(arg1=="A"&&arg2=="A")
    {
        return leve_1;
    }
    else if(arg1=="A"&&arg2=="B1")
    {
        return leve_1;
    }
    else if(arg1=="A"&&arg2=="B2")
    {
        return leve_2;
    }
    else if(arg1=="A"&&arg2=="C1")
    {
        return leve_2;
    }
    else if(arg1=="A"&&arg2=="C2")
    {
        return leve_3;
    }
    else if(arg1=="B"&&arg2=="A")
    {
        return leve_1;
    }
    else if(arg1=="B"&&arg2=="B1")
    {
        return leve_2;
    }
    else if(arg1=="B"&&arg2=="B2")
    {
        return leve_2;
    }
    else if(arg1=="B"&&arg2=="C1")
    {
        return leve_3;
    }
    else if(arg1=="B"&&arg2=="C2")
    {
        return leve_3;
    }
    else if(arg1=="C"&&arg2=="A")
    {
        return leve_1;
    }
    else if(arg1=="C"&&arg2=="B1")
    {
        return leve_2;
    }
    else if(arg1=="C"&&arg2=="B2")
    {
        return leve_3;
    }
    else if(arg1=="C"&&arg2=="C1")
    {
        return leve_3;
    }
    else if(arg1=="C"&&arg2=="C2")
    {
        return leve_4;
    }
    else if(arg1=="D"&&arg2=="A")
    {
        return leve_2;
    }
    else if(arg1=="D"&&arg2=="B1")
    {
        return leve_3;
    }
    else if(arg1=="D"&&arg2=="B2")
    {
        return leve_3;
    }
    else if(arg1=="D"&&arg2=="C1")
    {
        return leve_4;
    }
    else if(arg1=="D"&&arg2=="C2")
    {
        return leve_4;
    }
    else if(arg1=="E"&&arg2=="A")
    {
        return leve_3;
    }
    else if(arg1=="E"&&arg2=="B1")
    {
        return leve_3;
    }
    else if(arg1=="E"&&arg2=="B2")
    {
        return leve_4;
    }
    else if(arg1=="E"&&arg2=="C1")
    {
        return leve_4;
    }
    else if(arg1=="E"&&arg2=="C2")
    {
        return leve_4;
    }
}

//基础功能
void Rams::paintEvent(QPaintEvent *e)
{
    QStyleOption opt;
        opt.init(this);
        QPainter p(this);
        style()->drawPrimitive(QStyle::PE_Widget, &opt, &p, this);

}



//控件功能（可靠性）
void Rams::on_close_clicked()
{
    //关闭窗口
    this->close();
}

void Rams::on_maxmize_clicked()
{
    //最大化窗口
    this->showMaximized();
}

void Rams::on_minimize_clicked()
{
    //最小化窗口
    this->showMinimized();
}

void Rams::on_start_count_clicked()
{

    //查询底事件概率值
    QVariantList proList;
    QSqlQuery query;
    query.exec("select * from hlp");
    while(query.next())
    {
        proList<<query.value(1);
        //qDebug()<<query.value(1);
    }



    int n=proList.size();

    ui->rability_value->setText(QString::number(countRability(proList,n),10,10));
    ui->rability_value_2->setText(QString::number(1/(1-countRability(proList,n)),10,5));
    QString name = ui->Nodes ->currentText();
    ui->rability_value_3->setText(QString::number(countProImportance(name,proList,countRability(proList,n)),10,10));
    on_Pushbutton_Flag_1 = true;

}

void Rams::on_Nodes_activated(const QString &arg1)
{
    if(on_Pushbutton_Flag_1 == true)
    {QVariantList proList;
    QSqlQuery query;
    query.exec("select * from hlp");
    while(query.next())
    {
        proList<<query.value(1);
    }
    int n=proList.size();

    ui->rability_value_3->setText(QString::number(countProImportance(arg1,proList,countRability(proList,n)),10,10));
    }
}

void Rams::on_start_count_2_clicked()
{
    QString name =  ui->parts->currentText();
    //查询底事件概率值
    QList<double> proList;
    QSqlQuery query;
    query.exec("select * from hlp");
    while(query.next())
    {
        proList<<query.value(1).toDouble();
    }
    double p,t;
     if(name=="汇流排")
     {
         p=proList.at(0)+proList.at(1)+proList.at(2)+proList.at(3)+proList.at(4);
         t=1/p;
         ui->rability_value_4->setText(QString::number(p));
         ui->rability_value_5->setText(QString::number(t));
     }
     else if(name=="接触线")
     {
         p=proList.at(5)+proList.at(6)+proList.at(7)+proList.at(8)+proList.at(9);
         t=1/p;
         ui->rability_value_4->setText(QString::number(p));
         ui->rability_value_5->setText(QString::number(t));
     }
     else if(name=="绝缘子")
     {
         p=proList.at(10)+proList.at(11)+proList.at(12)+proList.at(13)+proList.at(14);
         t=1/p;
         ui->rability_value_4->setText(QString::number(p));
         ui->rability_value_5->setText(QString::number(t));
     }
     else if(name=="定位装置")
     {
         p=proList.at(15)+proList.at(16)+proList.at(2)+proList.at(13);
         t=1/p;
         ui->rability_value_4->setText(QString::number(p));
         ui->rability_value_5->setText(QString::number(t));
     }
     else if(name=="分段绝缘器")
     {
         p=proList.at(2)*2+proList.at(10)+proList.at(11)+proList.at(12)+proList.at(13)+proList.at(14)+proList.at(17);
         t=1/p;
         ui->rability_value_4->setText(QString::number(p));
         ui->rability_value_5->setText(QString::number(t));
     }
     else if(name=="上网引线")
     {
         p=proList.at(19)+proList.at(20);
         t=1/p;
         ui->rability_value_4->setText(QString::number(p));
         ui->rability_value_5->setText(QString::number(t));
     }
     on_Pushbutton_Flag_2 = true;
}

void Rams::on_parts_activated(const QString &arg1)
{
    if(on_Pushbutton_Flag_3)
    {
        //查询底事件概率值
        QList<double> proList;
        QSqlQuery query;
        query.exec("select * from hlp");
        while(query.next())
        {
            proList<<query.value(1).toDouble();
        }
        double p,t;
         if(arg1=="汇流排")
         {
             p=proList.at(0)+proList.at(1)+proList.at(2)+proList.at(3)+proList.at(4);
             t=1/p;
             ui->rability_value_4->setText(QString::number(p));
             ui->rability_value_5->setText(QString::number(t));
         }
         else if(arg1=="接触线")
         {
             p=proList.at(5)+proList.at(6)+proList.at(7)+proList.at(8)+proList.at(9);
             t=1/p;
             ui->rability_value_4->setText(QString::number(p));
             ui->rability_value_5->setText(QString::number(t));
         }
         else if(arg1=="绝缘子")
         {
             p=proList.at(10)+proList.at(11)+proList.at(12)+proList.at(13)+proList.at(14);
             t=1/p;
             ui->rability_value_4->setText(QString::number(p));
             ui->rability_value_5->setText(QString::number(t));
         }
         else if(arg1=="定位装置")
         {
             p=proList.at(15)+proList.at(16)+proList.at(2)+proList.at(13);
             t=1/p;
             ui->rability_value_4->setText(QString::number(p));
             ui->rability_value_5->setText(QString::number(t));
         }
         else if(arg1=="分段绝缘器")
         {
             p=proList.at(2)*2+proList.at(10)+proList.at(11)+proList.at(12)+proList.at(13)+proList.at(14)+proList.at(17);
             t=1/p;
             ui->rability_value_4->setText(QString::number(p));
             ui->rability_value_5->setText(QString::number(t));
         }
         else if(arg1=="上网引线")
         {
             p=proList.at(18)+proList.at(19);
             t=1/p;
             ui->rability_value_4->setText(QString::number(p));
             ui->rability_value_5->setText(QString::number(t));
         }
    }
}

void Rams::on_start_count_3_clicked()
{
    QString name = ui->parts_2->currentText();
    ui->rability_value_6->setText(QString::number(count_DisRe(name)));
    on_Pushbutton_Flag_3 = true;
}

void Rams::on_parts_2_activated(const QString &arg1)
{
    if(on_Pushbutton_Flag_3==true)
    {
      ui->rability_value_6->setText(QString::number(count_DisRe(arg1)));
    }
}


void Rams::on_Factor_activated(const QString &arg1)
{
    QSqlQuery query;
    //将节点输入到combobox中
    QSqlQueryModel *FactorModel = new QSqlQueryModel(this);
    if(arg1=="人员因素")
    {FactorModel->setQuery("select Human from human");
    ui->factors->setModel(FactorModel);

    int n = ui->factors->currentIndex();
    query.exec("select * from human");
    query.seek(n);
    ui->Safety_value->setText(query.value(1).toString());
    ui->Safety_value_2->setText(query.value(2).toString());
    QString PoValue = ui->Safety_value->text();
    QString SeValue = ui->Safety_value_2->text();
    ui->Safety_value_3->setText(count_Level(PoValue,SeValue));

    }
    else if(arg1=="设备因素")
    {FactorModel->setQuery("select Device from device");
    ui->factors->setModel(FactorModel);
    int n = ui->factors->currentIndex();
    query.exec("select * from device");
                 query.seek(n);
                 ui->Safety_value->setText(query.value(1).toString());
                 ui->Safety_value_2->setText(query.value(2).toString());
                 QString PoValue = ui->Safety_value->text();
                 QString SeValue = ui->Safety_value_2->text();
                 ui->Safety_value_3->setText(count_Level(PoValue,SeValue));
    }
    else if(arg1=="环境因素")
    {FactorModel->setQuery("select Environment from environment");
    ui->factors->setModel(FactorModel);
    int n = ui->factors->currentIndex();
    query.exec("select * from environment");
     query.seek(n);
     ui->Safety_value->setText(query.value(1).toString());
     ui->Safety_value_2->setText(query.value(2).toString());
     QString PoValue = ui->Safety_value->text();
     QString SeValue = ui->Safety_value_2->text();
     ui->Safety_value_3->setText(count_Level(PoValue,SeValue));
    }
    else if(arg1=="管理因素")
    {FactorModel->setQuery("select Management from management");
    ui->factors->setModel(FactorModel);
    int n = ui->factors->currentIndex();
    query.exec("select * from management");
      query.seek(n);
      ui->Safety_value->setText(query.value(1).toString());
      ui->Safety_value_2->setText(query.value(2).toString());
      QString PoValue = ui->Safety_value->text();
      QString SeValue = ui->Safety_value_2->text();
      ui->Safety_value_3->setText(count_Level(PoValue,SeValue));
    }
}

void Rams::on_factors_activated(const QString &arg1)
{
    QSqlQuery query;
    QString name = ui->Factor->currentText();

    int n = ui->factors->currentIndex();
            if(name=="人员因素"){
            query.exec("select * from human");
            query.seek(n);
            ui->Safety_value->setText(query.value(1).toString());
            ui->Safety_value_2->setText(query.value(2).toString());
            QString PoValue = ui->Safety_value->text();
            QString SeValue = ui->Safety_value_2->text();
            ui->Safety_value_3->setText(count_Level(PoValue,SeValue));
            }
      else if(name=="设备因素")
            {   query.exec("select * from device");
                query.seek(n);
                ui->Safety_value->setText(query.value(1).toString());
                ui->Safety_value_2->setText(query.value(2).toString());
                QString PoValue = ui->Safety_value->text();
                QString SeValue = ui->Safety_value_2->text();
                ui->Safety_value_3->setText(count_Level(PoValue,SeValue));
            }
            else if(name=="环境因素")
            {   query.exec("select * from environment");
                query.seek(n);
                ui->Safety_value->setText(query.value(1).toString());
                ui->Safety_value_2->setText(query.value(2).toString());
                QString PoValue = ui->Safety_value->text();
                QString SeValue = ui->Safety_value_2->text();
                ui->Safety_value_3->setText(count_Level(PoValue,SeValue));
            }
            else if(name=="管理因素")
            {   query.exec("select * from management");
                query.seek(n);
                ui->Safety_value->setText(query.value(1).toString());
                ui->Safety_value_2->setText(query.value(2).toString());
                QString PoValue = ui->Safety_value->text();
                QString SeValue = ui->Safety_value_2->text();
                ui->Safety_value_3->setText(count_Level(PoValue,SeValue));
            }
}

void Rams::on_changeButton_clicked()
{
    QSqlQuery query;
    QString name = ui->NODES ->currentText();
    double Pro = ui->NodePro->text().toDouble();
    query.exec(QString("update hlp set BoProbabilityValue=%1 where BottomEvent='%2'").arg(Pro).arg(name));
}

void Rams::on_PrintButton_clicked()
{
    QSqlQuery query;
    query.exec("select * from routput");
    query.first();
    mysheets->dynamicCall("Add");
    QAxObject *sheet = workbook->querySubObject("ActiveSheet");
    sheet -> setProperty("Name","可靠度输出报表");
    QAxObject *cell = sheet->querySubObject("Range(QVariant,QVariant)","A1");
    QString inStr = ui->lineEdit->text();
    QString inStr2 = ui->lineEdit_2->text();
    cell->dynamicCall("SetValue(const QVariant&)",QVariant("项目名称"));
    cell = sheet->querySubObject("Range(QVariant,QVariant)","B1");
    cell->dynamicCall("SetValue(const QVariant&)",QVariant("时间"));
    cell = sheet->querySubObject("Range(QVariant,QVariant)","C1");
    cell->dynamicCall("SetValue(const QVariant&)",QVariant("可靠度"));
    cell = sheet->querySubObject("Range(QVariant,QVariant)","D1");
    cell->dynamicCall("SetValue(const QVariant&)",QVariant("故障间隔时间"));
    cell = sheet->querySubObject("Range(QVariant,QVariant)","E1");
    cell->dynamicCall("SetValue(const QVariant&)",QVariant("可靠度分配"));
    cell = sheet->querySubObject("Range(QVariant,QVariant)","A2");
    cell->dynamicCall("SetValue(const QVariant&)",QVariant(inStr));
    cell = sheet->querySubObject("Range(QVariant,QVariant)","B2");
    cell->dynamicCall("SetValue(const QVariant&)",QVariant(inStr2));
    cell = sheet->querySubObject("Range(QVariant,QVariant)","C2");
    cell->dynamicCall("SetValue(const QVariant&)",QVariant(query.value("可靠值")));
    cell = sheet->querySubObject("Range(QVariant,QVariant)","D2");
    cell->dynamicCall("SetValue(const QVariant&)",QVariant(query.value("故障间隔时间")));
    cell = sheet->querySubObject("Range(QVariant,QVariant)","E2");
    cell->dynamicCall("SetValue(const QVariant&)",QVariant(query.value("可靠度分配")));
    workbook ->dynamicCall("SaveAs(const QString&)","D:\\Qt\\office\\可靠度输出报表.xls");
    workbook->dynamicCall("Close()");
    myexcel->dynamicCall("Quit()");
    QMessageBox::information(this,tr("完毕"),tr("Excel 工作表已保存！"));
    ui->PrintButton->setEnabled(false);
}
