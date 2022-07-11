
from openpyxl import Workbook, load_workbook

file_name=("matematik.xlsx")
yeni_workbook=Workbook()
yeni_worksheet=yeni_workbook.active
yeni_worksheet.title='Sonuçlar'


users_data={}


def kısı_ekle():
    kişi_input=input('kişi eklemek için ad girin')
    kişi_soyisim=input('kişi soyismini girin: ')
    kişi_numara=int(input('kişi numarasını girin: '))
    kişi_not=int(input('kişi notunu girin: '))
    users_data[kişi_input]={'soyisim':kişi_soyisim,'numara':kişi_numara}
    if kişi_not<=0 or kişi_not>=100:
        users_data[kişi_input]['sınavnotu']='invalid number'
    elif kişi_not>=85:
        users_data[kişi_input]['sınavnotu']=kişi_not       
        users_data[kişi_input]['durum']='geçti'
        users_data[kişi_input]['harf notu']='AA'
    elif kişi_not>=75:
        users_data[kişi_input]['sınavnotu']=kişi_not
        users_data[kişi_input]['durum']='geçti'
        users_data[kişi_input]['harf notu']='BB'
    elif kişi_not>=65:
        users_data[kişi_input]['sınavnotu']=kişi_not
        users_data[kişi_input]['durum']='geçti'
        users_data[kişi_input]['harf notu']='CC'
    elif kişi_not>=50:
        users_data[kişi_input]['sınavnotu']=kişi_not
        users_data[kişi_input]['durum']='koşullu geçti'
        users_data[kişi_input]['harf notu']='DC'
    else:
        users_data[kişi_input]['sınavnotu']=kişi_not
        users_data[kişi_input]['durum']='kaldı'
        users_data[kişi_input]['harf notu']='FF'
    

def invalid():
    print('Düzeltilecek kişi ismini girin\n')
    b=input('İsim: ')
    if b in users_data.keys():
        print(' Değeri girin\n')
        c=int(input('Değer: '))
        users_data[b]['sınavnotu']=c
    else:
        print('isim listesinde olmayan isim girdiniz')

while True:
    try:
        print('''ne  yapmak istersiniz ?
\t0= veri girişi
\t1=veri listesini göstermek
\t2=invalid  değer düzeltme
\t3=excele çıktı almak 
\tprogramı kapatmak için herhangi bir tuşa basın''')
        a=int(input())
        if a==0:
            try:
                kısı_ekle()
            except ValueError:
                print('yanlış  değer  girdiniz başa döndürülüyorsunuz')
        elif a==1:
            print(users_data)
        elif a==2:
            
            try:
                invalid()
            except(ValueError,KeyError):
                print('yanlış değer veya listede değer yok girdiniz başa döndürülüyorsunuz')
        elif a==3:
            print('excel')
            break
    except ValueError:
        break



