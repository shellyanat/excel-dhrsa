import streamlit as st

# For dealing with DataFrames
import pandas as pd
import numpy as np
import random

# For Download Result and Exce;
from io import BytesIO
import xlwt
from xlwt.Workbook import *
#from pyxlsb import open_workbook as open_xlsb

#page setting
st.set_page_config(layout="wide")
#page setting
st.markdown('''
<style>
/*center metric label*/
[data-testid="stMetricLabel"] > div:nth-child(1) {
    justify-content: center;
}

/*center metric value*/
[data-testid="stMetricValue"] > div:nth-child(1) {
    justify-content: center;
}
</style>
''', unsafe_allow_html=True)

st.markdown("""
<style>
div[data-testid="metric-container"] {
   background-color: rgba(28, 131, 225, 0.1);
   border: 1px solid rgba(28, 131, 225, 0.1);
   padding: 2% 2% 2% 8%;
   border-radius: 5px;
   color: rgb(30, 103, 119);
   overflow-wrap: break-word;
}

/* breakline for metric text         */
div[data-testid="metric-container"] > label[data-testid="stMetricLabel"] > div {
   overflow-wrap: break-word;
   white-space: break-spaces;
   color: blue;
}
</style>
"""
, unsafe_allow_html=True)

tabs_font_css = """
<style>
button[data-baseweb="tab"] {
  font-size: 28px;
}
</style>
"""

st.write(tabs_font_css, unsafe_allow_html=True)

'''
## Pengamanan Data pada File Excel
#### 1903685 Shellya Nur Atqiya
---
'''


#fpb
def FPB(m,n):
    if m<n:
        o=m
        m=n
        n=o
    s=m%n
    while s!=0:
        m=n
        n=s
        s=m%n
    return n

#cekprima
def cekprima(j):
    tes=0
    i=2
    while i <j:
        if FPB(j,i) == 1:
            tes=tes+0
        elif FPB(j,i)!=1:
            tes=tes+1
        i+=1
    return tes
        
#inversmodulo
def InvMod(a,b):
    inv=1
    while(inv*a)%b!=1:
        inv+=1
    return inv


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, header = False, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data

        

listTabs = [
    "Pembangkitan Kunci",
    "Enkripsi File",
    "Dekripsi File"]

whitespace = 35

tabs = st.tabs([s.center(whitespace,"\u2001") for s in listTabs])

with tabs[0]:
    with st.expander("LANGKAH-LANGKAH PENGGUNAAN"):
        '''
    Berikut langkah-langkah untuk melakukan pembangkitan kunci:

    1. Input dua nilai g, n, x, y dengan syarat:

    - Nilai g dan n harus bilangan prima yang besar
    
    - Nilai n > g > 126 (Nilai n lebih besar dari g dan 126)

    - Nilai x < n (Nilai x kurang dari n)

    - Nilai y < n (Nilai y kurang dari n)

    2. Klik tombol "Pembangkitan Kunci" agar program berjalan
    
    3. Simpan nilai kunci publik dan kunci privat dengan aman
        '''
    bag_g, bag_n = st.columns([5,5])
    with bag_g:
        g = st.text_input('Input nilai g :')
        if g:
            if cekprima(int(g)) > 0:
                st.write('nilai g harus merupakan bilangan prima, input nilai g baru')
    with bag_n:
        n = st.text_input('Input nilai n:')
        if n:
            if cekprima(int(n)) > 0:
                st.write('nilai n harus merupakan bilangan prima, input nilai n baru')
            if n <= g :
                st.write('nilai n harus lebih dari g, input nilai n baru')
    bag_x, bag_y = st.columns([5,5])
    with bag_x:
        x = st.text_input('Input nilai x :')
        if x:
            if int(x) >= int(n) :
                st.write('nilai x harus kurang dari nilai n, input nilai x baru')
    with bag_y:
        y = st.text_input('Input nilai y:')
        if y:
            if int(y) >= int(n) :
                st.write('nilai y harus kurang dari nilai n, input nilai y baru')
    if st.button('Bangkitkan Kunci'):
        g = int(g)
        n = int(n)
        x = int(x)
        y = int(y)
        X = (g**x)%n
        Y = (g**y)%n
        K1 = (X**y)%n
        K2 = (Y**x)%n
        if K1 == K2:
            cp_k = cekprima(K1)
            while cp_k!=0:
                K1 = K1+1
                cp_k = cekprima(K1)
            K2 = K1

            N1 = n*g*K1
            touN= (n-1)*(g-1)

            key_e = random.randrange(1,touN)
            z = FPB(key_e,touN)
            while z != 1:
                key_e = random.randrange(1,touN)
                z = FPB(key_e,touN)
            
            key_d = InvMod(key_e,touN)
            N2 = N1/K1
            N2 = int(N2)

            publickey = str(key_e) + ' ' + str(N1)
            privatkey = str(key_d) + ' ' + str(N2)

            col_1, col_2 = st.columns([5,5])
            with col_1:
                st.metric('Kunci Publik untuk Enkripsi', publickey)
            with col_2:
                with st.expander('Kunci Privat'):
                    st.metric('Kunci Privat untuk Dekripsi', privatkey)
             
with tabs[1]: 
    with st.expander("LANGKAH-LANGKAH PENGGUNAAN"):
        '''
    Berikut langkah-langkah untuk melakukan proses enkripsi:
    
    1. Input nilai kunci publik 
    
    2. Upload file yang akan dienkripsi dengan menekan tombol “Browse Files”

    3. File akan otomatis diproses untuk enkripsi, tunggu sampai proses enkripsi selesai
    
    4. Download file dengan menekan tombol “Download Encrypted File” untuk mendownload file yang telah dienkripsi.
        '''

    pubkey = st.text_input('Masukkan kunci publik:')

    uploaded_file = st.file_uploader("Upload File Excel:")
    if uploaded_file is not None:
   
        if st.button('Enkripsi File'):
            sheets = pd.ExcelFile(uploaded_file).sheet_names
                    
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')

            public_key = pubkey.split()
            e = public_key[0]
            N1 = public_key[1]

            for i, sheet in enumerate(sheets):
                dataframe = pd.read_excel(uploaded_file, sheet_name = sheet, header = None)
                dataframe = dataframe.astype(str)
                dataframe.fillna('nan-empty-values-kosong-dhrsa', inplace = True)
                for col in dataframe.columns:
                    for ind in dataframe.index:
                        a = dataframe[col][ind] #read per cell
                        a = str(a)
                        pjg_a = len(a) #panjang value
                        i = 0
                        cipherteks = ''
                        while i < pjg_a:
                            e = int(e)
                            N1 = int(N1)
                            m = a[i] #read per character
                            m = ord(m) #ubah character ke ascii
                            c = (m**e % N1)
                            if c < 100:
                                c = '0' + str(c)
                            c = str(c)
                            i = i+1
                            cipherteks = cipherteks + c + ' '
                        dataframe.loc[ind, col] = cipherteks 
                globals()['df' + str(i+1)] = dataframe.copy()

            #for i, sheet in enumerate(sheets):
                #workbook = writer.book
                wb = Workbook()
                worksheet = wb.add_sheet(sheet)
                globals()['df' + str(i+1)].to_excel(writer, header = False, index=False, sheet_name=sheet)
                #format1 = wb.add_format({'num_format': '0.00'}) 
                #worksheet.set_column('A:A', None, format1)  
            writer.save()
            processed_data = output.getvalue()


            data_xlsx = processed_data
            st.download_button(label='📥 Download Encrypted File',
                                data=data_xlsx,
                                file_name= 'encrypted_file.xlsx')
        if uploaded_file is None:
            st.write('')


with tabs[2]:
    with st.expander("LANGKAH-LANGKAH PENGGUNAAN"):
        '''
    Berikut langkah-langkah untuk melakukan proses dekripsi:

    1. Input nilai kunci privat 
    
    2. Upload file yang akan didekripsi dengan menekan tombol “Browse Files”
    
    3. File akan otomatis diproses untuk dekripsi, tunggu sampai proses dekripsi selesai
    
    4. Download file dengan menekan tombol “Download Decrypted File” untuk mendownload file yang telah didekripsi.    
        '''

    privkey = st.text_input('Masukkan kunci privat:')

    uploaded_file = st.file_uploader("Upload Encrypted File Excel:")
    if uploaded_file is not None:

        if st.button('Dekripsi File'):
            sheets = pd.ExcelFile(uploaded_file).sheet_names
                    
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')

            privat_key = privkey.split()
            d = privat_key[0]
            N2 = privat_key[1]

            for i, sheet in enumerate(sheets):
                dataframe = pd.read_excel(uploaded_file, sheet_name = sheet, header = None)
    
                dataframe.fillna('nan-empty-values-kosong-dhrsa', inplace = True)
                for col in dataframe.columns:
                    for ind in dataframe.index:
                        d = int(d)
                        N2 = int(N2)
                        a = dataframe[col][ind] #read per cell
                        b = a.split()
                        cipher = []
                        for z in b:
                            cipher.append(int(z))
                        plainteks = ''
                        for c in cipher:
                            m = (c**d % N2)
                            m = chr(m)
                            plainteks = plainteks + m
                        dataframe.loc[ind, col] = plainteks 
                        dataframe.replace("nan-empty-values-kosong-dhrsa", np.NaN, inplace=True)
                        
                globals()['df' + str(i+1)] = dataframe.copy()
                wb = Workbook()
                worksheet = wb.add_sheet(sheet)
                globals()['df' + str(i+1)].to_excel(writer, header = False, index=False, sheet_name=sheet)
                
            writer.save()
            processed_data = output.getvalue()
                
                
            data_xlsx = processed_data
            st.download_button(label='📥 Download Decrypted File',
                                data=data_xlsx,
                                file_name= 'decrypted_file.xlsx')



    if uploaded_file is None:
            st.write('')
