import os
import pandas as pd
from fpdf import FPDF


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    pdf = FPDF(orientation="P", unit="mm", format="A4")
        # включаем TTF шрифты, поддерживающие кириллицу
    pdf.add_font("Sans", style="", fname="C:/Fonts/NotoSans-Regular.ttf", uni=True)
    pdf.add_font("Sans", style="B", fname="C:/Fonts/NotoSans-Bold.ttf", uni=True)
    pdf.add_font("Sans", style="I", fname="C:/Fonts/NotoSans-Italic.ttf", uni=True)
    pdf.add_font("Sans", style="BI", fname="C:/Fonts/NotoSans-BoldItalic.ttf", uni=True)

    df = pd.read_excel('test.xlsx', index_col=0, sheet_name='Sheet1')
    pd.options.display.expand_frame_repr = False
    pd.set_option('display.max_columns', 6)
    pd.set_option('display.max_colwidth', 30)

    with open('1.txt', 'r', encoding="utf-8") as fp:
        txt1 = fp.read()
    with open('2.txt', 'r', encoding="utf-8") as fp:
        txt2 = fp.read()

    for i in range(0, len(df.index)):
        print(f"{i}из{len(df.index)}")
        if str(df.iloc[i]['Сыскин']) == 'нет' and str(df.iloc[i]['Управляющая компания']) != 'nan':
            adres = str(df.iloc[i]['Структурированный адрес'])
            upr_komp = str(df.iloc[i]['Управляющая компания'])
                # добавляем пустую страницу
            pdf.add_page()
                # задаем шрифт `Sans` ,
                # `Bold` (жирный) и размером 16

            pdf.image('Герб.jpg', x=100, y=None, w=10,  type='')
            pdf.set_font("Sans", "", 11)
            pdf.cell(190, 7, "КЕМЕРОВСКАЯ ОБЛАСТЬ-КУЗБАСС", align='C')
            pdf.ln(5)
            pdf.cell(190, 7, "НОВОКУЗНЕЦКИЙ ГОРОДСКОЙ ОКРУГ", align='C')
            pdf.ln(5)
            pdf.cell(190, 7, "АДМИНИСТРАЦИЯ ГОРОДА НОВОКУЗНЕЦКА", align='C')
            pdf.ln()
            pdf.cell(190, 7, '', border='T', ln=0, align='C')
            pdf.ln(1)
            pdf.cell(190, 7, '', border='T', ln=0, align='C')
            pdf.ln()

            pdf.set_font("Sans", "B", 13)
            pdf.cell(190, 7, "Уважаемый Собственник!", ln=20, align='C' )
            pdf.ln()

            pdf.set_font("Sans", "", 11)
            pdf.cell(190, 7, "Право собственности на принадлежащую Вам квартиру по адресу:")
            pdf.ln(12)

            pdf.set_font("Sans", "I", 11)
            pdf.cell(190, 7, adres, ln=20, align='C')
            pdf.ln()

            pdf.set_font("Sans", "", 11)
            pdf.cell(190, 7, "не зарегистрировано в Едином государственном реестре недвижимости (далее - ЕГРН).")
            pdf.ln(10)

            pdf.set_font("Sans", "", 11)
            pdf.multi_cell(190, 7, txt1)
            pdf.ln(10)

            pdf.set_font("Sans", "B", 12)
            pdf.multi_cell(190, 7, "С целью защиты Ваших прав и законных интересов Вам необходимо "
                             "зарегистрировать право собственности на принадлежащую Вам квартиру в ЕГРН!", align='C')
            pdf.ln(10)

            pdf.set_font("Sans", "", 11)
            pdf.multi_cell(190, 7, txt2)
            pdf.ln(2)

            pdf.set_font("Sans", "", 11)
            pdf.cell(190, 7, "По всем вопросам Вы можете обращаться по телефону:")
            pdf.ln()
            pdf.cell(190, 7, "8 (3843) 45-86-73, Новокузнецкий отдел Управления Росреестра по Кемеровской области")
            pdf.ln(20)

            pdf.cell(190, 7, "Администрация города Новокузнецка", align='R')
            pdf.ln()

            pdf.set_font("Sans", "I", 11)
            pdf.cell(190, 7, upr_komp, align='R')
            pdf.ln()

    pdf.output("Инком-С(доп).pdf")
    os.startfile("Инком-С(доп).pdf")
