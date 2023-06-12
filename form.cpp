#include "form.h"
#include "ui_form.h"
#include<QSqlQueryModel>
#include<Eigen/Dense>
#include<QSqlQuery>
#include<QMouseEvent>
#include<QDebug>
#include<QMovie>
#include<QProcess>
#include<QFileDialog>
#include<QTextStream>
using namespace Eigen;
bool Form::on_Pushbutton_Flag_1 = false;
bool Form::on_Pushbutton_Flag_2 = false;
bool Form::on_Pushbutton_Flag_3 = false;
bool Form::on_Pushbutton_Flag_4 = false;
bool Form::on_Pushbutton_Flag_5 = false;
bool Form::on_Pushbutton_Flag_6 = false;
Form::Form(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::Form)
{
    ui->setupUi(this);
    //去掉边框
    setWindowFlag(Qt::WindowType::FramelessWindowHint);

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
    ui->NODES->setModel(FtaModel);

    //将部件输入到parts
    QStringList list;
    list<<"汇流排"<<"接触线"<<"绝缘子"<<"定位装置"<<"分段绝缘器"<<"上网引线";
    ui->parts->addItems(list);
    //将文件名输入到filename
    QStringList filenamelist;
    filenamelist<<"耦合节点位移曲线"<<"弓头节点位移曲线"<<"接触压力曲线";
    ui->filename->addItems(filenamelist);

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
    ui->yinsuji->addItems(strList);

    //将因素输入combobox_2中
    QSqlQueryModel *RaDModel = new QSqlQueryModel(this);
    RaDModel->setQuery("select Factors from mohu");
    ui->yinsuji_2->setModel(RaDModel);
    ui->beizeji->setModel(RaDModel);


    //将因素输入Factor中
    QStringList nameList;
    nameList<<"人员因素"<<"设备因素"<<"环境因素"<<"管理因素";
    ui->Factor->addItems(nameList);

    //将具体因素输入factors中
    QSqlQueryModel *FactorModel = new QSqlQueryModel(this);
    FactorModel->setQuery("select Human from human");
    ui->factors->setModel(FactorModel);

    //提示
    //ui->Nodes->lineEdit()->setPlaceholderText("节点名称");

    //默认显示第一个窗口
    ui->stackedWidget->setCurrentIndex(0);
    ui->stackedWidget_2->setCurrentIndex(0);


    //将子系统故障输入到SubSystemFaultComboBox
    QStringList SubSystemFault;
    SubSystemFault<<"汇流排失效"<<"接触线失效"<<"绝缘子失效"<<"定位装置失效"<<"分段绝缘器失效"<<"上网引线失效";
    ui->SubSystemFaultComboBox->addItems(SubSystemFault);

    //安全初始化
    QStringList NameList;
    NameList<<"人员因素"<<"设备因素"<<"环境因素"<<"管理因素";
    ui->main_factors->addItems(NameList);

    QSqlQueryModel *FactorModelS = new QSqlQueryModel(this);
    FactorModelS->setQuery("select Human from human");
    ui->Factors->setModel(FactorModel);

    //绘制表格
    QSqlQueryModel *Pmodel = new QSqlQueryModel;
    QSqlQueryModel *Smodel = new QSqlQueryModel;
    Pmodel->setQuery("select 可能性等级 as 可能性等级, 等级说明 as 等级说明, 发生情况 as 发生情况 from poss");
    Smodel->setQuery("select 严重等级 as 严重等级, 等级说明 as 等级说明, 事故后果说明 as 事故后果说明 from severity");

    ui->tableView->setModel(Pmodel);
    ui->tableView_2->setModel(Smodel);
    ui->tableView->setColumnWidth(0,120);
    ui->tableView->setColumnWidth(1,120);
    ui->tableView->setColumnWidth(2,190);
    ui->tableView->setRowHeight(0,80);
    ui->tableView->setRowHeight(1,80);
    ui->tableView->setRowHeight(2,80);
    ui->tableView->setRowHeight(3,80);
    ui->tableView->setRowHeight(4,80);
    ui->tableView->setRowHeight(5,80);
    ui->tableView_2->setColumnWidth(0,120);
    ui->tableView_2->setColumnWidth(1,120);
    ui->tableView_2->setColumnWidth(2,500);
    ui->tableView_2->setRowHeight(0,77);
    ui->tableView_2->setRowHeight(1,77);
    ui->tableView_2->setRowHeight(2,77);
    ui->tableView_2->setRowHeight(3,77);
    ui->tableView_2->setRowHeight(4,77);
    ui->tableView_2->setRowHeight(5,77);

    //添加质量
    QStringList m_listComboBox;
    m_listComboBox<<"M1"<<"M2"<<"M3"<<"MEQ";
    ui->m_ComboBox->addItems(m_listComboBox);
    //添加刚度
    QStringList k_listComboBox;
    k_listComboBox<<"K"<<"K1"<<"K2"<<"K3"<<"KEQ"<<"KEQI";
    ui->k_ComboBox->addItems(k_listComboBox);
    //添加阻尼
    QStringList c_listComboBox;
    c_listComboBox<<"C1"<<"C2"<<"C3";
    ui->c_ComboBox->addItems(c_listComboBox);
}

Form::~Form()
{
    delete ui;
}
//计算函数
double Form::countRability(QVariantList list, int n)
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

double Form::countProImportance(QString NodeName, QVariantList list, double r)
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

float Form::count_DisRe(QString name)
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
    float Rs=0.98,r=0;
    QString rD= QString("%1/%2/%3/%4").arg(QString::number(pow(Rs,F(0,0)))).arg(QString::number(pow(Rs,F(0,1)))).arg(QString::number(pow(Rs,F(0,2)))).arg(QString::number(pow(Rs,F(0,3))));
    query.exec(QString("update routput set 可靠度分配='%1'").arg(rD));
    if(name=="接触悬挂")
    {
        r=pow(Rs,F(0,0));
    }
    else if(name=="支持定位")
    {
        r=pow(Rs,F(0,1));
    }
    else if(name=="绝缘装置")
    {
        r=pow(Rs,F(0,2));
    }
    else if(name=="附加导线")
    {
        r=pow(Rs,F(0,3));
    }
    return r;
}

QString Form::count_Level(QString arg1, QString arg2)
{
    QString leve_1="不可接受",leve_2="可接受",leve_3="可忽略";

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
        return leve_1;
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
        return leve_1;
    }
    else if(arg1=="B"&&arg2=="B2")
    {
        return leve_2;
    }
    else if(arg1=="B"&&arg2=="C1")
    {
        return leve_2;
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
        return leve_2;
    }
    else if(arg1=="C"&&arg2=="C1")
    {
        return leve_2;
    }
    else if(arg1=="C"&&arg2=="C2")
    {
        return leve_3;
    }
    else if(arg1=="D"&&arg2=="A")
    {
        return leve_2;
    }
    else if(arg1=="D"&&arg2=="B1")
    {
        return leve_2;
    }
    else if(arg1=="D"&&arg2=="B2")
    {
        return leve_2;
    }
    else if(arg1=="D"&&arg2=="C1")
    {
        return leve_3;
    }
    else if(arg1=="D"&&arg2=="C2")
    {
        return leve_3;
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
        return leve_3;
    }
    else if(arg1=="E"&&arg2=="C1")
    {
        return leve_3;
    }
    else if(arg1=="E"&&arg2=="C2")
    {
        return leve_3;
    }
    else{
        return 0;
    }
}

void Form::onOut()
{
    //实时回显数据
    qDebug()<<ansys->readAllStandardOutput().data();
}

void Form::getExcelContent(QVector<QVector<QString> > &map)
{

    exc_res.clear();
    QAxObject *excel = NULL;
    QAxObject *workbooks = NULL;
    QAxObject *workbook = NULL;
    //QString filePath = QFileDialog::getOpenFileName(
    //            this, QStringLiteral("选择Excel文件"),"",
    //            QStringLiteral("Excel file(*.csv)"));
    QString filePath;
    if(ui->filename->currentText()=="耦合节点位移曲线")
    {filePath = "F:\\AnsysWork\\DispY.csv"; }
    else if(ui->filename->currentText()=="弓头节点位移曲线")
    {filePath = "F:\\AnsysWork\\DispY1.csv"; }
    else if(ui->filename->currentText()=="接触压力曲线")
    {filePath = "F:\\AnsysWork\\DispY2.csv"; }

    if(filePath.isEmpty())return;

    CoInitializeEx(NULL, COINIT_MULTITHREADED);
    excel = new QAxObject("Excel.Application");
    if(!excel)
    {
        qDebug()<<"EXCEL 对象丢失！";
    }

    workbooks = excel->querySubObject("Workbooks");//所有excel文件
    if(0==workbooks)
    {
        qDebug()<<"Kong000000000";
        return;
    }
    workbook = workbooks->querySubObject("Open (const QString &)",filePath);//按路径获取文件
    if(0==workbook)
    {
        qDebug()<<"workbook000000000";
        return;
    }
    QAxObject *worksheet = workbook->querySubObject("WorkSheets(int)", 1);//读取第一个表

    QAxObject *usedRange = worksheet->querySubObject("UsedRange");//有数据的矩形区域
    QAxObject * rows = usedRange->querySubObject("Rows");
    QAxObject * columns = usedRange->querySubObject("Columns");
    intRows = rows->property("Count").toInt();
    intCols = columns->property("Count").toInt();
    //qDebug()<<"行数："<<intRows;
    //qDebug()<<"列数："<<intCols;

    QVariant var = usedRange->dynamicCall("Value");
    foreach (QVariant varRow, var.toList())
    {
        QVector<QString> vecDataRow;
        foreach(QVariant var, varRow.toList())
        {
            vecDataRow.push_back(var.toString());
        }
        map.push_back(vecDataRow);
    }
    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    if(excel)
    {
        delete excel;
        excel = NULL;
    }
}

//基础功能设置
void Form::paintEvent(QPaintEvent *e)
{
    QStyleOption opt;
        opt.init(this);
        QPainter p(this);
        style()->drawPrimitive(QStyle::PE_Widget, &opt, &p, this);
}
void Form::mousePressEvent(QMouseEvent *event)
{
    auto e = static_cast<QMouseEvent*>(event);
    if (e->button() == Qt::LeftButton) {
        m_isDrag = true;
        m_offPos = e->globalPos()-this->frameGeometry().topLeft();
    }
    event->accept();
}

void Form::mouseMoveEvent(QMouseEvent *event)
{
    auto e = static_cast<QMouseEvent*>(event);
    if(m_isDrag)
    {
    this->move(e->globalPos()-m_offPos);

    }
    event->accept();
}

void Form::mouseReleaseEvent(QMouseEvent *event)
{
    auto e = static_cast<QMouseEvent*>(event);
    if(e->button()==Qt::LeftButton)
    {
        m_isDrag = false;
    }
}

void Form::mouseDoubleClickEvent(QMouseEvent *event)
{
    if(event->button()==Qt::LeftButton)
    {
        if(windowState()!=Qt::WindowFullScreen)
            setWindowState(Qt::WindowFullScreen);
        else setWindowState(Qt::WindowNoState);
    }
}


//控件功能设置
void Form::on_PrintButton_clicked()
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

void Form::on_close_clicked()
{
    //关闭窗口
    this->close();
}

void Form::on_maxmize_clicked()
{
    //最大化窗口
    this->showMaximized();
}

void Form::on_minimize_clicked()
{
    //最小化窗口
    this->showMinimized();
}

void Form::on_start_count_clicked()
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

void Form::on_Nodes_activated(const QString &arg1)
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

void Form::on_start_count_2_clicked()
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

void Form::on_parts_activated(const QString &arg1)
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

void Form::on_start_count_3_clicked()
{
    QString name = ui->parts_2->currentText();
    ui->rability_value_6->setText(QString::number(count_DisRe(name)));
    on_Pushbutton_Flag_3 = true;
}

void Form::on_parts_2_activated(const QString &arg1)
{
    if(on_Pushbutton_Flag_3==true)
    {
      ui->rability_value_6->setText(QString::number(count_DisRe(arg1)));
    }
}

void Form::on_Factor_activated(const QString &arg1)
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

void Form::on_factors_activated(const QString &arg1)
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

void Form::on_changeButton_clicked()
{
    QSqlQuery query;
    QString name = ui->NODES ->currentText();
    double Pro = ui->NodePro->text().toDouble();
    query.exec(QString("update hlp set BoProbabilityValue=%1 where BottomEvent='%2'").arg(Pro).arg(name));
}

void Form::on_listWidget_currentRowChanged(int currentRow)
{
    ui->stackedWidget->setCurrentIndex(currentRow);
}

void Form::on_NODES_activated(const QString &arg1)
{
    QSqlQuery query;
    query.exec("select * from hlp");
    int n = ui->NODES->currentIndex();
    query.seek(n);
    ui->NodePro->setText(query.value(1).toString());
}

void Form::on_changeButtonE_clicked()
{
    QSqlQuery query;

    float a1 = ui->expert_01_input->text().toFloat();
    float a2 = ui->expert_02_input->text().toFloat();
    float a3 = ui->expert_03_input->text().toFloat();

    QString name = ui->yinsuji->currentText();
    QString name2 = ui->beizeji->currentText();
    if(name=="接触悬挂")
    {
    query.exec(QString("update mohu set JS_1=%1 where Factors='%2'").arg(a1).arg(name2));
    query.exec(QString("update mohu set JS_2=%1 where Factors='%2'").arg(a2).arg(name2));
    query.exec(QString("update mohu set JS_3=%1 where Factors='%2'").arg(a3).arg(name2));
    }
    else if(name=="支持定位"){
        query.exec(QString("update mohu set ZS_1=%1 where Factors='%2'").arg(a1).arg(name2));
        query.exec(QString("update mohu set ZS_2=%1 where Factors='%2'").arg(a2).arg(name2));
        query.exec(QString("update mohu set ZS_3=%1 where Factors='%2'").arg(a3).arg(name2));
    }
    else if(name=="绝缘装置"){
        query.exec(QString("update mohu set jS_11=%1 where Factors='%2'").arg(a1).arg(name2));
        query.exec(QString("update mohu set jS_22=%1 where Factors='%2'").arg(a2).arg(name2));
        query.exec(QString("update mohu set jS_33=%1 where Factors='%2'").arg(a3).arg(name2));
    }
    else if(name=="附加导线"){
        query.exec(QString("update mohu set FS_1=%1 where Factors='%2'").arg(a1).arg(name2));
        query.exec(QString("update mohu set FS_2=%1 where Factors='%2'").arg(a2).arg(name2));
        query.exec(QString("update mohu set FS_3=%1 where Factors='%2'").arg(a3).arg(name2));
    }
}

void Form::on_yinsuji_activated(const QString &arg1)
{
    QSqlQuery query;
    query.exec("select * from mohu");
    QStringList list,list1,list2,list3,list4,list5,list6,list7,list8,list9,list10,list11;
    while(query.next())
    {
       list<<query.value("JS_1").toString();
       list1<<query.value("JS_2").toString();
       list2<<query.value("JS_3").toString();
       list3<<query.value("ZS_1").toString();
       list4<<query.value("ZS_2").toString();
       list5<<query.value("ZS_3").toString();
       list6<<query.value("jS_11").toString();
       list7<<query.value("jS_22").toString();
       list8<<query.value("jS_33").toString();
       list9<<query.value("FS_1").toString();
       list10<<query.value("FS_2").toString();
       list11<<query.value("FS_3").toString();
    }
    QString name = ui->beizeji->currentText();
    if(arg1=="接触悬挂"&&name=="重要程度")
        {
        ui->expert_01_input->setText(list.at(0));
        ui->expert_02_input->setText(list1.at(0));
        ui->expert_03_input->setText(list2.at(0));
    }
    else if(arg1=="接触悬挂"&&name=="复杂度")
        {
        ui->expert_01_input->setText(list.at(1));
        ui->expert_02_input->setText(list1.at(1));
        ui->expert_03_input->setText(list2.at(1));
    }
    else if(arg1=="接触悬挂"&&name=="成本度")
        {
        ui->expert_01_input->setText(list.at(2));
        ui->expert_02_input->setText(list1.at(2));
        ui->expert_03_input->setText(list2.at(2));
    }
    else if(arg1=="接触悬挂"&&name=="维修因素")
        {
        ui->expert_01_input->setText(list.at(3));
        ui->expert_02_input->setText(list1.at(3));
        ui->expert_03_input->setText(list2.at(3));
    }
    else if(arg1=="接触悬挂"&&name=="工作环境")
        {
        ui->expert_01_input->setText(list.at(4));
        ui->expert_02_input->setText(list1.at(4));
        ui->expert_03_input->setText(list2.at(4));
    }
    else if(arg1=="接触悬挂"&&name=="技术水平")
        {
        ui->expert_01_input->setText(list.at(5));
        ui->expert_02_input->setText(list1.at(5));
        ui->expert_03_input->setText(list2.at(5));
    }
    else if(arg1=="支持定位"&&name=="重要程度")
        {
        ui->expert_01_input->setText(list3.at(0));
        ui->expert_02_input->setText(list4.at(0));
        ui->expert_03_input->setText(list5.at(0));
    }
    else if(arg1=="支持定位"&&name=="复杂度")
        {
        ui->expert_01_input->setText(list3.at(1));
        ui->expert_02_input->setText(list4.at(1));
        ui->expert_03_input->setText(list5.at(1));
    }
    else if(arg1=="支持定位"&&name=="成本")
        {
        ui->expert_01_input->setText(list3.at(2));
        ui->expert_02_input->setText(list4.at(2));
        ui->expert_03_input->setText(list5.at(2));
    }
    else if(arg1=="支持定位"&&name=="维修因素")
        {
        ui->expert_01_input->setText(list3.at(3));
        ui->expert_02_input->setText(list4.at(3));
        ui->expert_03_input->setText(list5.at(3));
    }
    else if(arg1=="支持定位"&&name=="工作环境")
        {
        ui->expert_01_input->setText(list3.at(4));
        ui->expert_02_input->setText(list4.at(4));
        ui->expert_03_input->setText(list5.at(4));
    }
    else if(arg1=="支持定位"&&name=="技术水平")
        {
        ui->expert_01_input->setText(list3.at(5));
        ui->expert_02_input->setText(list4.at(5));
        ui->expert_03_input->setText(list5.at(5));
    }
    else if(arg1=="绝缘装置"&&name=="重要程度")
        {
        ui->expert_01_input->setText(list6.at(0));
        ui->expert_02_input->setText(list7.at(0));
        ui->expert_03_input->setText(list8.at(0));
    }
    else if(arg1=="绝缘装置"&&name=="复杂度")
        {
        ui->expert_01_input->setText(list6.at(1));
        ui->expert_02_input->setText(list7.at(1));
        ui->expert_03_input->setText(list8.at(1));
    }
    else if(arg1=="绝缘装置"&&name=="成本")
        {
        ui->expert_01_input->setText(list6.at(2));
        ui->expert_02_input->setText(list7.at(2));
        ui->expert_03_input->setText(list8.at(2));
    }
    else if(arg1=="绝缘装置"&&name=="维修因素")
        {
        ui->expert_01_input->setText(list6.at(3));
        ui->expert_02_input->setText(list7.at(3));
        ui->expert_03_input->setText(list8.at(3));
    }
    else if(arg1=="绝缘装置"&&name=="工作环境")
        {
        ui->expert_01_input->setText(list6.at(4));
        ui->expert_02_input->setText(list7.at(4));
        ui->expert_03_input->setText(list8.at(4));
    }
    else if(arg1=="绝缘装置"&&name=="技术水平")
        {
        ui->expert_01_input->setText(list6.at(5));
        ui->expert_02_input->setText(list7.at(5));
        ui->expert_03_input->setText(list8.at(5));
    }
    else if(arg1=="附加导线"&&name=="重要程度")
        {
        ui->expert_01_input->setText(list9.at(0));
        ui->expert_02_input->setText(list10.at(0));
        ui->expert_03_input->setText(list11.at(0));
    }
    else if(arg1=="附加导线"&&name=="复杂度")
        {
        ui->expert_01_input->setText(list9.at(1));
        ui->expert_02_input->setText(list10.at(1));
        ui->expert_03_input->setText(list11.at(1));
    }
    else if(arg1=="附加导线"&&name=="成本")
        {
        ui->expert_01_input->setText(list9.at(2));
        ui->expert_02_input->setText(list10.at(2));
        ui->expert_03_input->setText(list11.at(2));
    }
    else if(arg1=="附加导线"&&name=="维修因素")
        {
        ui->expert_01_input->setText(list9.at(3));
        ui->expert_02_input->setText(list10.at(3));
        ui->expert_03_input->setText(list11.at(3));
    }
    else if(arg1=="附加导线"&&name=="工作环境")
        {
        ui->expert_01_input->setText(list9.at(4));
        ui->expert_02_input->setText(list10.at(4));
        ui->expert_03_input->setText(list11.at(4));
    }
    else if(arg1=="附加导线"&&name=="技术水平")
        {
        ui->expert_01_input->setText(list9.at(5));
        ui->expert_02_input->setText(list10.at(5));
        ui->expert_03_input->setText(list11.at(5));}
}

void Form::on_changeButtonE_2_clicked()
{
    QSqlQuery query;
    QString name = ui->yinsuji_2->currentText();
    float Exe = ui->expert_01_input_2->text().toFloat();
    float Exe_2 = ui->expert_02_input_2->text().toFloat();
    float Exe_3 = ui->expert_03_input_2->text().toFloat();
    query.exec(QString("update mohu set FScore_1=%1 where Factors='%2'").arg(Exe).arg(name));
    query.exec(QString("update mohu set FScore_2=%1 where Factors='%2'").arg(Exe_2).arg(name));
    query.exec(QString("update mohu set FScore_3=%1 where Factors='%2'").arg(Exe_3).arg(name));
}

void Form::on_yinsuji_2_activated(const QString &arg1)
{
    QList<float> list,list_1,list_2;
    QSqlQuery query;
    query.exec("select * from mohu");
    while(query.next())
    {
       list<<query.value(5).toFloat();
       list_1<<query.value(6).toFloat();
       list_2<<query.value(7).toFloat();
    }
    if(arg1=="重要程度")
    {
      ui->expert_01_input_2->setText(QString::number(list.at(0)));
      ui->expert_02_input_2->setText(QString::number(list_1.at(0)));
      ui->expert_03_input_2->setText(QString::number(list_2.at(0)));
    }
    else if(arg1=="复杂度")
    {
      ui->expert_01_input_2->setText(QString::number(list.at(1)));
      ui->expert_02_input_2->setText(QString::number(list_1.at(1)));
      ui->expert_03_input_2->setText(QString::number(list_2.at(1)));
    }
    else if(arg1=="成本")
    {
      ui->expert_01_input_2->setText(QString::number(list.at(2)));
      ui->expert_02_input_2->setText(QString::number(list_1.at(2)));
      ui->expert_03_input_2->setText(QString::number(list_2.at(2)));
    }
    else if(arg1=="维修因素")
    {
      ui->expert_01_input_2->setText(QString::number(list.at(3)));
      ui->expert_02_input_2->setText(QString::number(list_1.at(3)));
      ui->expert_03_input_2->setText(QString::number(list_2.at(3)));
    }
    else if(arg1=="工作环境")
    {
      ui->expert_01_input_2->setText(QString::number(list.at(4)));
      ui->expert_02_input_2->setText(QString::number(list_1.at(4)));
      ui->expert_03_input_2->setText(QString::number(list_2.at(4)));
    }
    else if(arg1=="技术水平")
    {
      ui->expert_01_input_2->setText(QString::number(list.at(5)));
      ui->expert_02_input_2->setText(QString::number(list_1.at(5)));
      ui->expert_03_input_2->setText(QString::number(list_2.at(5)));
    }
}

void Form::on_factorChangeButton_clicked()
{
    QSqlQuery query;
    QString name = ui->Factors_input->text();
    QString arg1 = ui->main_factors->currentText();
    QString arg2 = ui->Factors->currentText();
    if(arg1=="人员因素"){
    query.exec(QString("update human set Human='%1' where Human='%2'").arg(name).arg(arg2));
    }
    else if(arg1=="设备因素")
    {    query.exec(QString("update device set Device='%1' where Device='%2'").arg(name).arg(arg2));
    }
    else if(arg1=="环境因素")
    {    query.exec(QString("update environment set Environment='%1' where Environment='%2'").arg(name).arg(arg2));
    }
    else if(arg1=="管理因素")
    {    query.exec(QString("update management set Management='%1' where Management='%2'").arg(name).arg(arg2));

    }
}

void Form::on_changeButton_all_clicked()
{
    QString name = ui->main_factors->currentText();
    QString name_2 = ui->Factors->currentText();
    QString PoValue = ui->properity_input->text();
    QString SeValue = ui->serious_input->text();
    QSqlQuery query;
    if(name=="人员因素"){
    query.exec(QString("update human set Possibility='%1' where Human='%2'").arg(PoValue).arg(name_2));
    query.exec(QString("update human set Harm='%1' where Human='%2'").arg(SeValue).arg(name_2));
    query.exec(QString("update human set RiskLevel='%1' where Human='%2'").arg(count_Level(PoValue,SeValue)).arg(name_2));

    }
    else if(name=="设备因素")
    {    query.exec(QString("update device set Possibility='%1' where Device='%2'").arg(PoValue).arg(name_2));
         query.exec(QString("update device set Harm='%1' where Device='%2'").arg(SeValue).arg(name_2));
         query.exec(QString("update human set RiskLevel='%1' where Device='%2'").arg(count_Level(PoValue,SeValue)).arg(name_2));

    }
    else if(name=="环境因素")
    {    query.exec(QString("update environment set Possibility='%1' where Environment='%2'").arg(PoValue).arg(name_2));
         query.exec(QString("update environment set Harm='%1' where Environment='%2'").arg(PoValue).arg(name_2));
         query.exec(QString("update human set RiskLevel='%1' where Environment='%2'").arg(count_Level(PoValue,SeValue)).arg(name_2));

    }
    else if(name=="管理因素")
    {    query.exec(QString("update management set Possibility='%1' where Management='%2'").arg(PoValue).arg(name_2));
         query.exec(QString("update management set Harm='%1' where Management='%2'").arg(PoValue).arg(name_2));
         query.exec(QString("update human set RiskLevel='%1' where Management='%2'").arg(count_Level(PoValue,SeValue)).arg(name_2));

    }
}

void Form::on_main_factors_activated(const QString &arg1)
{
    QSqlQuery query;
    //将节点输入到combobox中
    QSqlQueryModel *FactorModel = new QSqlQueryModel(this);
    if(arg1=="人员因素")
    {FactorModel->setQuery("select Human from human");
    ui->Factors->setModel(FactorModel);

    int n = ui->Factors->currentIndex();
    query.exec("select * from human");
    query.seek(n);
    ui->properity_input->setText(query.value(1).toString());
    ui->serious_input->setText(query.value(2).toString());
    QString PoValue = ui->properity_input->text();
    QString SeValue = ui->serious_input->text();
    }
    else if(arg1=="设备因素")
    {FactorModel->setQuery("select Device from device");
    ui->Factors->setModel(FactorModel);
    int n = ui->Factors->currentIndex();
    query.exec("select * from device");
                 query.seek(n);
                 ui->properity_input->setText(query.value(1).toString());
                 ui->serious_input->setText(query.value(2).toString());
                 QString PoValue = ui->properity_input->text();
                 QString SeValue = ui->serious_input->text();
                 ui->label_4->setText(count_Level(PoValue,SeValue));
    }
    else if(arg1=="环境因素")
    {FactorModel->setQuery("select Environment from environment");
    ui->Factors->setModel(FactorModel);
    int n = ui->Factors->currentIndex();
    query.exec("select * from environment");
     query.seek(n);
     ui->properity_input->setText(query.value(1).toString());
     ui->serious_input->setText(query.value(2).toString());
     QString PoValue = ui->properity_input->text();
     QString SeValue = ui->serious_input->text();
     ui->label_4->setText(count_Level(PoValue,SeValue));
    }
    else if(arg1=="管理因素")
    {FactorModel->setQuery("select Management from management");
    ui->Factors->setModel(FactorModel);
    int n = ui->Factors->currentIndex();
    query.exec("select * from management");
      query.seek(n);
      ui->properity_input->setText(query.value(1).toString());
      ui->serious_input->setText(query.value(2).toString());
      QString PoValue = ui->properity_input->text();
      QString SeValue = ui->serious_input->text();
      ui->label_4->setText(count_Level(PoValue,SeValue));
    }
}

void Form::on_Factors_activated(const QString &arg1)
{
    QSqlQuery query;
    ui->Factors_input->setText(arg1);
    QString name = ui->main_factors->currentText();

    int n = ui->Factors->currentIndex();
            if(name=="人员因素"){
            query.exec("select * from human");
            query.seek(n);
            ui->properity_input->setText(query.value(1).toString());
            ui->serious_input->setText(query.value(2).toString());
            QString PoValue = ui->properity_input->text();
            QString SeValue = ui->serious_input->text();
            }
      else if(name=="设备因素")
            {   query.exec("select * from device");
                query.seek(n);
                ui->properity_input->setText(query.value(1).toString());
                ui->serious_input->setText(query.value(2).toString());
                QString PoValue = ui->properity_input->text();
                QString SeValue = ui->serious_input->text();
            }
            else if(name=="环境因素")
            {   query.exec("select * from environment");
                query.seek(n);
                ui->properity_input->setText(query.value(1).toString());
                ui->serious_input->setText(query.value(2).toString());
                QString PoValue = ui->properity_input->text();
                QString SeValue = ui->serious_input->text();
            }
            else if(name=="管理因素")
            {   query.exec("select * from management");
                query.seek(n);
                ui->properity_input->setText(query.value(1).toString());
                ui->serious_input->setText(query.value(2).toString());
                QString PoValue = ui->properity_input->text();
                QString SeValue = ui->serious_input->text();
            }
}

void Form::on_SubSystemFaultComboBox_activated(const QString &arg1)
{

    if (arg1=="汇流排失效")
    {
        QPixmap pixmap("bottomEvent01.png");
        pixmap=pixmap.scaled(800,600,Qt::KeepAspectRatio);
        ui->SubSystemFaultView->setPixmap(pixmap);
    }
    else if (arg1=="接触线失效")
    {
        QPixmap pixmap("bottomEvent02.png");
        pixmap=pixmap.scaled(800,600,Qt::KeepAspectRatio);
        ui->SubSystemFaultView->setPixmap(pixmap);
    }
    else if (arg1=="绝缘子失效")
    {
        QPixmap pixmap("bottomEvent03.png");
        pixmap=pixmap.scaled(800,600,Qt::KeepAspectRatio);
        ui->SubSystemFaultView->setPixmap(pixmap);
    }
    else if (arg1=="定位装置失效")
    {
        QPixmap pixmap("bottomEvent04.png");
        pixmap=pixmap.scaled(800,600,Qt::KeepAspectRatio);
        ui->SubSystemFaultView->setPixmap(pixmap);
    }
    else if (arg1=="分段绝缘器失效")
    {
        QPixmap pixmap("bottomEvent05.png");
        pixmap=pixmap.scaled(800,600,Qt::KeepAspectRatio);
        ui->SubSystemFaultView->setPixmap(pixmap);
    }
    else if (arg1=="上网引线失效")
    {
        QPixmap pixmap("bottomEvent06.png");
        pixmap=pixmap.scaled(800,600,Qt::KeepAspectRatio);
        ui->SubSystemFaultView->setPixmap(pixmap);
    }
}

void Form::on_SystemSetList_currentRowChanged(int currentRow)
{
        ui->stackedWidget_2->setCurrentIndex(currentRow);
}

void Form::on_GenerateButton_clicked()
{
 //QMovie *movie = new QMovie (this);
  //  movie ->setFileName("ansys.gif");
   // movie->start();//播放
  //  ui->ansysScene->setMovie(movie);

    //调用ANSYS
    ansys = new QProcess;
    ui->progressBar1->setValue(0);
    QString program("F:\\AnsysWork\\ansys.bat");
   // QStringList arguements;
   // arguements <<"-p ansys"<<"-b"<<"-dir \"F:\\AnsysWork\""<<"-j \"file\""<<"-i \"F:\\AnsysWork\\file.dat\""<<"-o \"F:\\AnsysWork\\file.txt\"";
    ansys->setProcessChannelMode(QProcess::MergedChannels);
    connect(ansys,&QProcess::readyReadStandardOutput,this,&Form::onOut);
    ansys->start(program);
    if(!ansys->waitForStarted())
    {
        qDebug()<<"start failed:"<<ansys->errorString();
    }
    else{
        qDebug()<<"start success:";
        ui->progressBar1->setValue(15);
    }
    int i=25;
    while (!ansys->waitForFinished())
    {

        qDebug()<<"finish failed:"<<ansys->errorString();
        if(i<100)
        {ui->progressBar1->setValue(i);
        i+=25;}
    }
        ui->progressBar1->setValue(100);
        qDebug()<<"finish success:";

}

void Form::on_beizeji_activated(const QString &name)
{
    QSqlQuery query;
    query.exec("select * from mohu");
    QStringList list,list1,list2,list3,list4,list5,list6,list7,list8,list9,list10,list11;
    while(query.next())
    {
       list<<query.value("JS_1").toString();
       list1<<query.value("JS_2").toString();
       list2<<query.value("JS_3").toString();
       list3<<query.value("ZS_1").toString();
       list4<<query.value("ZS_2").toString();
       list5<<query.value("ZS_3").toString();
       list6<<query.value("jS_11").toString();
       list7<<query.value("jS_22").toString();
       list8<<query.value("jS_33").toString();
       list9<<query.value("FS_1").toString();
       list10<<query.value("FS_2").toString();
       list11<<query.value("FS_3").toString();
    }
    QString arg1 = ui->yinsuji->currentText();
    if(arg1=="接触悬挂"&&name=="重要程度")
        {
        ui->expert_01_input->setText(list.at(0));
        ui->expert_02_input->setText(list1.at(0));
        ui->expert_03_input->setText(list2.at(0));
    }
    else if(arg1=="接触悬挂"&&name=="复杂度")
        {
        ui->expert_01_input->setText(list.at(1));
        ui->expert_02_input->setText(list1.at(1));
        ui->expert_03_input->setText(list2.at(1));
    }
    else if(arg1=="接触悬挂"&&name=="成本度")
        {
        ui->expert_01_input->setText(list.at(2));
        ui->expert_02_input->setText(list1.at(2));
        ui->expert_03_input->setText(list2.at(2));
    }
    else if(arg1=="接触悬挂"&&name=="维修因素")
        {
        ui->expert_01_input->setText(list.at(3));
        ui->expert_02_input->setText(list1.at(3));
        ui->expert_03_input->setText(list2.at(3));
    }
    else if(arg1=="接触悬挂"&&name=="工作环境")
        {
        ui->expert_01_input->setText(list.at(4));
        ui->expert_02_input->setText(list1.at(4));
        ui->expert_03_input->setText(list2.at(4));
    }
    else if(arg1=="接触悬挂"&&name=="技术水平")
        {
        ui->expert_01_input->setText(list.at(5));
        ui->expert_02_input->setText(list1.at(5));
        ui->expert_03_input->setText(list2.at(5));
    }
    else if(arg1=="支持定位"&&name=="重要程度")
        {
        ui->expert_01_input->setText(list3.at(0));
        ui->expert_02_input->setText(list4.at(0));
        ui->expert_03_input->setText(list5.at(0));
    }
    else if(arg1=="支持定位"&&name=="复杂度")
        {
        ui->expert_01_input->setText(list3.at(1));
        ui->expert_02_input->setText(list4.at(1));
        ui->expert_03_input->setText(list5.at(1));
    }
    else if(arg1=="支持定位"&&name=="成本")
        {
        ui->expert_01_input->setText(list3.at(2));
        ui->expert_02_input->setText(list4.at(2));
        ui->expert_03_input->setText(list5.at(2));
    }
    else if(arg1=="支持定位"&&name=="维修因素")
        {
        ui->expert_01_input->setText(list3.at(3));
        ui->expert_02_input->setText(list4.at(3));
        ui->expert_03_input->setText(list5.at(3));
    }
    else if(arg1=="支持定位"&&name=="工作环境")
        {
        ui->expert_01_input->setText(list3.at(4));
        ui->expert_02_input->setText(list4.at(4));
        ui->expert_03_input->setText(list5.at(4));
    }
    else if(arg1=="支持定位"&&name=="技术水平")
        {
        ui->expert_01_input->setText(list3.at(5));
        ui->expert_02_input->setText(list4.at(5));
        ui->expert_03_input->setText(list5.at(5));
    }
    else if(arg1=="绝缘装置"&&name=="重要程度")
        {
        ui->expert_01_input->setText(list6.at(0));
        ui->expert_02_input->setText(list7.at(0));
        ui->expert_03_input->setText(list8.at(0));
    }
    else if(arg1=="绝缘装置"&&name=="复杂度")
        {
        ui->expert_01_input->setText(list6.at(1));
        ui->expert_02_input->setText(list7.at(1));
        ui->expert_03_input->setText(list8.at(1));
    }
    else if(arg1=="绝缘装置"&&name=="成本")
        {
        ui->expert_01_input->setText(list6.at(2));
        ui->expert_02_input->setText(list7.at(2));
        ui->expert_03_input->setText(list8.at(2));
    }
    else if(arg1=="绝缘装置"&&name=="维修因素")
        {
        ui->expert_01_input->setText(list6.at(3));
        ui->expert_02_input->setText(list7.at(3));
        ui->expert_03_input->setText(list8.at(3));
    }
    else if(arg1=="绝缘装置"&&name=="工作环境")
        {
        ui->expert_01_input->setText(list6.at(4));
        ui->expert_02_input->setText(list7.at(4));
        ui->expert_03_input->setText(list8.at(4));
    }
    else if(arg1=="绝缘装置"&&name=="技术水平")
        {
        ui->expert_01_input->setText(list6.at(5));
        ui->expert_02_input->setText(list7.at(5));
        ui->expert_03_input->setText(list8.at(5));
    }
    else if(arg1=="附加导线"&&name=="重要程度")
        {
        ui->expert_01_input->setText(list9.at(0));
        ui->expert_02_input->setText(list10.at(0));
        ui->expert_03_input->setText(list11.at(0));
    }
    else if(arg1=="附加导线"&&name=="复杂度")
        {
        ui->expert_01_input->setText(list9.at(1));
        ui->expert_02_input->setText(list10.at(1));
        ui->expert_03_input->setText(list11.at(1));
    }
    else if(arg1=="附加导线"&&name=="成本")
        {
        ui->expert_01_input->setText(list9.at(2));
        ui->expert_02_input->setText(list10.at(2));
        ui->expert_03_input->setText(list11.at(2));
    }
    else if(arg1=="附加导线"&&name=="维修因素")
        {
        ui->expert_01_input->setText(list9.at(3));
        ui->expert_02_input->setText(list10.at(3));
        ui->expert_03_input->setText(list11.at(3));
    }
    else if(arg1=="附加导线"&&name=="工作环境")
        {
        ui->expert_01_input->setText(list9.at(4));
        ui->expert_02_input->setText(list10.at(4));
        ui->expert_03_input->setText(list11.at(4));
    }
    else if(arg1=="附加导线"&&name=="技术水平")
        {
        ui->expert_01_input->setText(list9.at(5));
        ui->expert_02_input->setText(list10.at(5));
        ui->expert_03_input->setText(list11.at(5));
    }
}

void Form::on_ViewButton_clicked()
{

    fileName = QFileDialog::getOpenFileName(this);
    QFile file(fileName);
        if(file.open(QIODevice::ReadOnly|QIODevice::Text))
        {
            QTextStream stream(&file);
            strAll=stream.readAll();
            ui->Batch_path_LineEdit->setText(strAll);
            replace_Name=strAll;
        }
    file.close();
    on_Pushbutton_Flag_4 = true;
}

void Form::on_ChangeButton_2_clicked()
{
    QFile writeFile(fileName);
    if(on_Pushbutton_Flag_4){
    if(writeFile.open(QIODevice::WriteOnly|QIODevice::Text))
    {
        QTextStream stream(&writeFile);
        strList=strAll.split("\n");
        for (int i=0;i<strList.count();i++) {
            if(strList.at(i).contains(strAll))
            {
                QString tempStr=strList.at(i);
                tempStr.replace(0,tempStr.length(),replace_Name);
                stream<<tempStr;
           }
            else
            {
                stream<<strList.at(i);
            }
        }
    }
    }
    writeFile.close();

    QFile writeFile_1(fileName_1);
    if(on_Pushbutton_Flag_5)
    {
        QString v_value="VI="+ui->v_value_LineEdit->text();
        QString l_value="LSP="+ui->lsp_value_LineEdit->text();
        QString f_value="FI="+ui->F0_value_LineEdit->text();
        QString p_value="F_POINT="+ui->p_value_LineEdit->text();
        strList_1.replace(11,l_value);
        strList_1.replace(12,v_value);
        strList_1.replace(15,m1_value);
        strList_1.replace(16,m2_value);
        strList_1.replace(17,m3_value);
        strList_1.replace(18,k_value);
        strList_1.replace(19,k1_value);
        strList_1.replace(20,k2_value);
        strList_1.replace(21,k3_value);
        strList_1.replace(22,c1_value);
        strList_1.replace(23,c2_value);
        strList_1.replace(24,c3_value);
        strList_1.replace(25,f_value);
        strList_1.replace(26,meq_value);
        strList_1.replace(27,keq_value);
        strList_1.replace(28,keqi_value);
        strList_1.replace(30,p_value);
    if(writeFile_1.open(QIODevice::WriteOnly|QIODevice::Text))
    {
        QTextStream stream(&writeFile_1);
        for (int i=0;i<strList_1.count();i++)
        {
              stream<<strList_1.at(i)<<'\n';
        }
    }
    }
   writeFile_1.close();
}

void Form::on_Batch_path_LineEdit_textChanged(const QString &arg1)
{
    strAll=replace_Name;
    replace_Name=arg1;
}

void Form::on_ViewButton_6_clicked()
{
    fileName_1 = QFileDialog::getOpenFileName(this);
    QFile file(fileName_1);
        if(file.open(QIODevice::ReadOnly|QIODevice::Text))
        {
            QTextStream stream(&file);
            QString tempStr;
            strAll_1=stream.readAll();
            strList_1=strAll_1.split("\n");
            tempStr = V_parseDate();
            ui->v_value_LineEdit->setText(V_parseDate());
            ui->lsp_value_LineEdit->setText(L_parseDate());
            ui->F0_value_LineEdit->setText(F_parseDate());
            ui->p_value_LineEdit->setText(P_parseDate());
            ui->m_value_LineEdit->setText(M_parseDate());
            ui->k_value_LineEdit->setText(K_parseDate());
            ui->c_value_LineEdit->setText(C_parseDate());
            replace_Name_1=strList_1;
        }
    file.close();
    on_Pushbutton_Flag_5 = true;
m1_value=strList_1.at(15);
m2_value=strList_1.at(16);
m3_value=strList_1.at(17);
meq_value=strList_1.at(26);
k1_value=strList_1.at(19);
k2_value=strList_1.at(20);
k3_value=strList_1.at(21);
keq_value=strList_1.at(27);
keqi_value=strList_1.at(28);
k_value=strList_1.at(18);
c1_value=strList_1.at(22);
c2_value=strList_1.at(23);
c3_value=strList_1.at(24);
}

QString Form::V_parseDate()
{
    char array[]={'V','I','='};
    int length = sizeof(array)/sizeof (char);
    QString tempStr=strList_1.at(12);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(12).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
    return tempStr;
}

QString Form::L_parseDate()
{
    char array[]={'L','S','P','='};
    int length = sizeof(array)/sizeof (char);
    QString tempStr=strList_1.at(11);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(11).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
    return tempStr;
}

QString Form::F_parseDate()
{
    char array[]={'F','I','='};
    int length = sizeof(array)/sizeof (char);
    QString tempStr=strList_1.at(25);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(25).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
    return tempStr;
}

QString Form::P_parseDate()
{
    char array[]={'F','_','P','O','I','N','T','='};
    int length = sizeof(array)/sizeof (char);
    QString tempStr=strList_1.at(30);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(30).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
    return tempStr;
}

QString Form::M_parseDate()
{

    char array[]={'M','I','E','Q','='};
    int length = sizeof(array)/sizeof (char);
    QString Name=ui->m_ComboBox->currentText();
    if(Name=="M1")
    {
    QString tempStr=strList_1.at(15);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(15).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }
    else if(Name=="M2")
    {
    QString tempStr=strList_1.at(16);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(16).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }

    else if(Name=="M3")
    {
    QString tempStr=strList_1.at(17);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(17).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }
    else if(Name=="MEQ")
    {
    QString tempStr=strList_1.at(26);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(26).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }

}

QString Form::K_parseDate()
{
    char array[]={'K','I','V','Q','='};
    int length = sizeof(array)/sizeof (char);
    QString Name=ui->k_ComboBox->currentText();
    if(Name=="K")
    {
    QString tempStr=strList_1.at(18);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(18).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }
    else if(Name=="K1")
    {
    QString tempStr=strList_1.at(19);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(19).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }

    else if(Name=="K2")
    {
    QString tempStr=strList_1.at(20);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(20).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }
    else if(Name=="K3")
    {
    QString tempStr=strList_1.at(21);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(21).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }
    else if(Name=="KEQ")
    {
    QString tempStr=strList_1.at(27);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(27).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }
    else if(Name=="KEQI")
    {
    QString tempStr=strList_1.at(28);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(28).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }
}

QString Form::C_parseDate()
{
    char array[]={'C','I','='};
    int length = sizeof(array)/sizeof (char);
    QString Name=ui->c_ComboBox->currentText();
    if(Name=="C1")
    {
    QString tempStr=strList_1.at(22);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(22).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }
    else if(Name=="C2")
    {
    QString tempStr=strList_1.at(23);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(23).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }

    else if(Name=="C3")
    {
    QString tempStr=strList_1.at(24);
    for (int i=0;i<length;i++) {
        QString tmp= QString(array[i]);
        if(strList_1.at(24).contains(tmp))
        {
        tempStr=tempStr.replace(tmp,"");
        }
    }
        return tempStr;
    }


}

void Form::on_m_ComboBox_activated(const QString &arg1)
{
    QString change_mlineEdit;
    change_mlineEdit=M_parseDate();
    ui->m_value_LineEdit->setText(change_mlineEdit);
}

void Form::on_k_ComboBox_activated(const QString &arg1)
{
    QString change_klineEdit;
    change_klineEdit=K_parseDate();
    ui->k_value_LineEdit->setText(change_klineEdit);
}

void Form::on_c_ComboBox_activated(const QString &arg1)
{
    QString change_clineEdit;
    change_clineEdit=C_parseDate();
    ui->c_value_LineEdit->setText(change_clineEdit);
}

void Form::on_m_value_LineEdit_textChanged(const QString &arg1)
{
    QString Name=ui->m_ComboBox->currentText();
    if(Name=="M1")
    {
    m1_value="MI="+arg1;
    }
    else if(Name=="M2")
    {
    m2_value="MII="+arg1;
    }
    else if(Name=="M3")
    {
    m3_value="MIII="+arg1;
    }
    else if(Name=="MEQ")
    {
    meq_value="MQ="+arg1;
    }

}

void Form::on_k_value_LineEdit_textChanged(const QString &arg1)
{
    QString Name=ui->k_ComboBox->currentText();
    if(Name=="K")
    {
    k_value="KIV="+arg1;
    }
    else if(Name=="K1")
    {
    k1_value="KI="+arg1;
    }

    else if(Name=="K2")
    {
    k2_value="KII="+arg1;
    }
    else if(Name=="K3")
    {
    k3_value="KIII="+arg1;
    }
    else if(Name=="KEQ")
    {
    keq_value="KQ="+arg1;
    }
    else if(Name=="KEQI")
    {
    keqi_value="KQI="+arg1;
    }
}

void Form::on_c_value_LineEdit_textChanged(const QString &arg1)
{
    QString Name=ui->c_ComboBox->currentText();
    if(Name=="C1")
    {
    c1_value="CI="+arg1;
    }
    else if(Name=="C2")
    {
    c2_value="CII="+arg1;
    }
    else if(Name=="C3")
    {
    c3_value="CIII="+arg1;
    }

}

void Form::on_GenerateButton_2_clicked()
{
    getExcelContent(exc_res);
    //mChart->close();
           //qDebug()<<"intCols----"<<intCols;
           //qDebug()<<"intRows----"<<intRows;
               if (on_Pushbutton_Flag_6==true)
               {
                   mChart->removeSeries(splineSeries);//删除上一个曲线
                   mChart = new QChart();
                       splineSeries = new QSplineSeries();  //QSplineSeries 平滑曲线  QLineSeries折线
                       QVector<QPointF> points;
                       for (int i=0; i<intRows; ++i)
                       {
                           points.append(QPointF(exc_res[i][0].toDouble(),exc_res[i][1].toDouble()));
                       }
                       splineSeries->replace(points);
                       mChart->addSeries(splineSeries);//添加新的折线图
               }
               else{
                   mChart = new QChart();
                   splineSeries = new QSplineSeries ;
               for(int i=0; i<intRows; ++i)
               {

                   splineSeries->append(exc_res[i][0].toDouble(),exc_res[i][1].toDouble());
               }
               mChart->addSeries(splineSeries);}

           mChart->setTheme(QChart::ChartThemeLight);
           mChart->legend()->hide();
           if(ui->filename->currentText()=="耦合节点位移曲线")
           {mChart->setTitle("耦合节点位移曲线"); }
           else if(ui->filename->currentText()=="弓头节点位移曲线")
           {mChart->setTitle("弓头节点位移曲线"); }
           else if(ui->filename->currentText()=="接触压力曲线")
           {mChart->setTitle("接触压力曲线"); }
           // 设置标题
           mChart->createDefaultAxes();                // 基于已添加到图表中的series为图表创建轴。以前添加到图表中的任何轴都将被删除。
           //mChart->axes(Qt::Vertical).first()->setRange(0, 1);  // 设置Y轴的范围

           ui->widget->setRenderHint(QPainter::Antialiasing);  // 设置抗锯齿
           ui->widget->setChart(mChart);//拖入界面的控件名为widget
           on_Pushbutton_Flag_6 = true;

}
