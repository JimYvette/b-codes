#ifndef FORM_H
#define FORM_H

#include <QWidget>
#include<QChartView>
#include<QSplineSeries>
#include<QLineSeries>
#include<QFileDialog>
#include<QList>
#include<QDebug>
#include<QMessageBox>
#include<QAxObject>
#include<QPainter>
#include<Eigen/Dense>
#include<QProcess>
#include<QtCharts>
QT_CHARTS_USE_NAMESPACE
namespace Ui {
class Form;
}

class Form : public QWidget
{
    Q_OBJECT
    static bool on_Pushbutton_Flag_1;
    static bool on_Pushbutton_Flag_2;
    static bool on_Pushbutton_Flag_3;
    static bool on_Pushbutton_Flag_4;
    static bool on_Pushbutton_Flag_5;
        static bool on_Pushbutton_Flag_6;
public:
    explicit Form(QWidget *parent = nullptr);
    ~Form();

    QVector<QVector<QString>> exc_res;
    int intRows;
    int intCols;

private slots:
    void on_PrintButton_clicked();
    double countRability(QVariantList list,int n);
    double countProImportance(QString NodeName,QVariantList list,double r);
    float count_DisRe(QString name);
    QString count_Level(QString arg1,QString arg2);
    void onOut();
    void getExcelContent(QVector<QVector<QString>>& result);


    void paintEvent(QPaintEvent *e);

    void mousePressEvent(QMouseEvent *event);

    void mouseMoveEvent(QMouseEvent *event);

    void mouseReleaseEvent(QMouseEvent *event);

    void mouseDoubleClickEvent(QMouseEvent *event);


    void on_close_clicked();

    void on_maxmize_clicked();

    void on_minimize_clicked();

    void on_start_count_clicked();

    void on_Nodes_activated(const QString &arg1);

    void on_start_count_2_clicked();

    void on_parts_activated(const QString &arg1);

    void on_start_count_3_clicked();

    void on_parts_2_activated(const QString &arg1);

    void on_Factor_activated(const QString &arg1);

    void on_factors_activated(const QString &arg1);

    void on_changeButton_clicked();

    void on_listWidget_currentRowChanged(int currentRow);


    void on_NODES_activated(const QString &arg1);

    void on_changeButtonE_clicked();

    void on_yinsuji_activated(const QString &arg1);

    void on_changeButtonE_2_clicked();

    void on_yinsuji_2_activated(const QString &arg1);

    void on_factorChangeButton_clicked();

    void on_changeButton_all_clicked();

    void on_main_factors_activated(const QString &arg1);

    void on_Factors_activated(const QString &arg1);

    void on_SubSystemFaultComboBox_activated(const QString &arg1);

    void on_SystemSetList_currentRowChanged(int currentRow);

    void on_GenerateButton_clicked();

    void on_beizeji_activated(const QString &name);

    void on_ViewButton_clicked();

    void on_ChangeButton_2_clicked();

    void on_Batch_path_LineEdit_textChanged(const QString &arg1);

    void on_ViewButton_6_clicked();

    QString V_parseDate();
    QString L_parseDate();
    QString F_parseDate();
    QString P_parseDate();
    QString M_parseDate();
    QString K_parseDate();
    QString C_parseDate();

    void on_m_ComboBox_activated(const QString &arg1);

    void on_k_ComboBox_activated(const QString &arg1);

    void on_c_ComboBox_activated(const QString &arg1);

    void on_m_value_LineEdit_textChanged(const QString &arg1);

    void on_k_value_LineEdit_textChanged(const QString &arg1);

    void on_c_value_LineEdit_textChanged(const QString &arg1);

    void on_GenerateButton_2_clicked();

private:
    Ui::Form *ui;
    QAxObject *myexcel;
    QAxObject *myworks;
    QAxObject *workbook;
    QAxObject *mysheets;
    QPoint m_offPos;
    QProcess *ansys;
    bool m_isDrag=false;
    QString fileName;
    QString strAll;
    QStringList strList;
    QString replace_Name;
    QString fileName_1;
    QString strAll_1;
    QStringList strList_1;
    QStringList replace_Name_1;
    QString m1_value;
    QString m2_value;
    QString m3_value;
    QString meq_value;
    QString k1_value;
    QString k2_value;
    QString k3_value;
    QString keq_value;
    QString keqi_value;
    QString k_value;
    QString c1_value;
    QString c2_value;
    QString c3_value;
    QChart *mChart;
    QSplineSeries *splineSeries;
};

#endif // FORM_H
