# -*- coding: utf-8 -*-

file1 = r"c:\Users\Khanh\Downloads\Book1.xlsx"

import openpyxl
import pyodbc

def put_nt1():
    book = openpyxl.load_workbook(file1)
    sheet = book.active

    for row in sheet.iter_rows(values_only=True):
        if row[0] != 'MST':
            # conn = pyodbc.connect('DRIVER={SQL Server};SERVER=DESKTOP-G635SF6;DATABASE=HR;TRUSTED_CONNECTION=yes;')
            conn = pyodbc.connect("Driver={SQL Server}; Server=172.16.60.100; Database=HR; UID=huynguyen; PWD=Namthuan@123;")
            cursor = conn.cursor()
            
            mst = row[0]
            thechamcong = int(row[1])
            hoten = row[2]
            sdt = row[3]
            ngaysinh = f"'{row[4]}'" if row[4] else 'NULL'
            gioitinh = row[5]
            cccd = row[6]
            ngaycapcccd = f"'{row[7]}'" if row[7] else 'NULL'
            noicapcccd = row[8]
            cmt = 'NULL'
            thuongtru = row[10]
            thonxom = row[11]
            phuongxa = row[12]
            quanhuyen = row[13]
            thanhpho = row[14]
            dantoc = f"N'{row[15]}'"
            quoctich = row[16]
            tongiao = row[17]
            trinhdo = row[18]
            noisinh = row[19]
            tamtru = row[20]
            sobhxh = row[21]
            masothue = row[22]
            nganhang = row[23]
            sotaikhoan = row[24]
            connho = row[25]
            tencon1 = f"'{row[26]}'" if row[26] else 'NULL'
            tencon2 = f"'{row[28]}'" if row[27] else 'NULL'
            tencon3 = f"'{row[30]}'" if row[28] else 'NULL'
            tencon4 = f"'{row[32]}'" if row[29] else 'NULL'
            tencon5 = f"'{row[34]}'" if row[30] else 'NULL'
            ngaysinhcon1 = f"'{row[27]}'" if row[31] else 'NULL'
            ngaysinhcon2 = f"'{row[29]}'" if row[32] else 'NULL'
            ngaysinhcon3 = f"'{row[31]}'" if row[33] else 'NULL'
            ngaysinhcon4 = f"'{row[33]}'" if row[34] else 'NULL'
            ngaysinhcon5 = f"'{row[35]}'" if row[35] else 'NULL'
            anh = row[36]
            nguoithan = row[37]
            sdtnguoithan = row[38]
            loaihopdong = row[39]
            ngaykyhd = f"'{row[40]}'" if row[40] else 'NULL'
            ngayhethanhd = f"'{row[41]}'" if row[41] else 'NULL'
            jobtitlevn = row[42]
            hccategory = row[43]
            gradecode = row[44]
            factory = row[45]
            department = row[46]
            chucvu = row[47]
            sectioncode = row[48]
            sectiondescription = row[49]
            line = row[50]
            emp_type = row[51]
            jobtitleen = row[52]
            positioncode = row[53]
            positiondescription = row[54]
            luongcoban = int(row[55]) if row[55] else 0
            phucap = row[56]
            tienphucap = int(row[57]) if row[57] else 0
            ngayvao = row[58] if row[58] else 'NULL'
            ngaynghi = f"'{row[59]}'" if row[59] else 'NULL'
            trangthai = row[60]
            ngayvaonoithamnien = f"'{row[61]}'" if row[61] else 'NULL'
            matkhau = row[62]
            ngaykihdtv = f"'{row[63]}'" if row[63] else 'NULL' 
            ngayhethanhdtv = f"'{row[64]}'" if row[64] else 'NULL'
            ngaykihdcthl1 = f"'{row[65]}'" if row[65] else 'NULL'
            ngayhethanhdcthl1 = f"'{row[66]}'" if row[66] else 'NULL'
            ngaykihdcthl2 = f"'{row[67]}'" if row[67] else 'NULL'
            ngayhethanhdcthl2 = f"'{row[68]}'" if row[68] else 'NULL'
            ngaykihdvthl = f"'{row[69]}'" if row[69] else 'NULL'
            truongbophan = f"'{row[70]}'"
            ghichu = row[71]
            query=f"""
                    INSERT INTO HR.dbo.Danh_sach_CBCNV
                    VALUES
                    (
                        '{mst}'
                        ,'{thechamcong}'
                        ,N'{hoten}'
                        ,'{sdt}'
                        ,{ngaysinh}
                        ,N'{gioitinh}'
                        ,'{cccd}'
                        ,{ngaycapcccd}
                        ,N'{noicapcccd}'
                        ,N'{cmt}'
                        ,N'{thuongtru}'
                        ,N'{thonxom}'
                        ,N'{phuongxa}'
                        ,N'{quanhuyen}'
                        ,N'{thanhpho}'
                        ,{dantoc}
                        ,N'{quoctich}'
                        ,N'{tongiao}'
                        ,N'{trinhdo}'
                        ,N'{noisinh}'
                        ,N'{tamtru}'
                        ,N'{sobhxh}'
                        ,N'{masothue}'
                        ,N'{nganhang}'
                        ,N'{sotaikhoan}'
                        ,N'{connho}'
                        ,{tencon1}
                        ,{ngaysinhcon1}
                        ,{tencon2}
                        ,{ngaysinhcon2}
                        ,{tencon3}
                        ,{ngaysinhcon3}
                        ,{tencon4}
                        ,{ngaysinhcon4}
                        ,{tencon5}
                        ,{ngaysinhcon5}
                        ,N'{anh}'
                        ,N'{nguoithan}'
                        ,'{sdtnguoithan}'
                        ,N'{loaihopdong}'
                        ,{ngaykyhd}
                        ,{ngayhethanhd}
                        ,N'{jobtitlevn}'
                        ,N'{hccategory}'
                        ,N'{gradecode}'
                        ,N'{factory}'
                        ,N'{department}'
                        ,N'{chucvu}'
                        ,N'{sectioncode}'
                        ,N'{sectiondescription}'
                        ,N'{line}'
                        ,N'{emp_type}'
                        ,N'{jobtitleen}'
                        ,N'{positioncode}'
                        ,N'{positiondescription}'
                        ,'{luongcoban}'
                        ,N'{phucap}'
                        ,'{tienphucap}'
                        ,'{ngayvao}'
                        ,{ngaynghi}
                        ,N'{trangthai}'
                        ,{ngayvaonoithamnien}
                        ,'{matkhau}'
                        ,{ngaykihdtv}
                        ,{ngayhethanhdtv}
                        ,{ngaykihdcthl1}
                        ,{ngayhethanhdcthl1}
                        ,{ngaykihdcthl2}
                        ,{ngayhethanhdcthl2}
                        ,{ngaykihdvthl}
                        ,'N'
                        ,N'{ghichu}'
                    )
                """
            print(mst)
            if factory == 'NT2' and int(mst)>2965:
                try:
                    print(cursor.execute(query))
                    conn.commit()
                    conn.close()
                except Exception as e:
                    print(query)
                    print(e)
                    conn.close()
                    break            
        
# def put_nt2():
#     book = openpyxl.load_workbook(file2)
#     sheet = book.active

#     for row in sheet.iter_rows(values_only=True):
#         conn = pyodbc.connect('DRIVER={SQL Server};SERVER=DESKTOP-G635SF6;DATABASE=HR;TRUSTED_CONNECTION=yes;')
#         cursor = conn.cursor()
#         if "STT" not in str(row[0]):
#             stt= str(row[0])
#             masothe = str(row[1])
#             thechamcong = str(row[2])
#             hoten = str(row[3])
#             gioitinh = str(row[4])
#             dantoc = str(row[5])
#             tongiao = str(row[6])
#             ngaysinh = row[7]
#             namsinh = str(row[8])
#             quequan = str(row[9])
#             thuongtru = str(row[10])
#             thanhpho = str(row[11])
#             quanhuyen = str(row[12])
#             phuongxa = str(row[13])
#             thonxom = str(row[14])
#             cccd = str(row[15])
#             ngaycapcccd = row[16]
#             noicapcccd = str(row[17])
#             tamtru = str(row[18])
#             trinhdo = str(row[19])
#             chuyenmon = str(row[20])
#             masothue = str(row[21])
#             sotaikhoan = str(row[22])
#             nganhang = str(row[23])
#             chinhanhnganhang = str(row[24])
#             sodienthoai = str(row[25])
#             sodienthoainguoithan = str(row[26])
#             ngayhocviec = row[27]
#             ngaythuviec = row[28]
#             ngaychinhthuc = row[29]
#             ngaynghiviec = row[30]
#             hccategory = str(row[31])
#             factory = str(row[32])
#             department = str(row[33])
#             sectioncode = str(row[34])
#             section = str(row[35])
#             line = str(row[36])
#             employeetype = str(row[37])
#             positioncode = str(row[38])
#             positioncodedescription = str(row[39])
#             gradecode = str(row[40])
#             chucvu = str(row[41])
#             jobtitlevn = str(row[42])
#             jobtitleen = str(row[43])
#             tinhtranglamviec = str(row[44])
#             hopdong = str(row[45])
#             if factory == "NT2":
#                 query=f"""
#                     INSERT INTO HR.dbo.Danh_sach_CBCNV
#                     VALUES
#                     (
#                         '{masothe}'
#                         ,'{thechamcong}'
#                         ,N'{hoten}'
#                         ,'{sodienthoai}'
#                         ,'{ngaysinh}'
#                         ,N'{gioitinh}'
#                         ,'{cccd}'
#                         ,'{ngaycapcccd}'
#                         ,N'{noicapcccd}'
#                         ,NULL
#                         ,N'{thuongtru}'
#                         ,N'{thonxom}'
#                         ,N'{phuongxa}'
#                         ,N'{quanhuyen}'
#                         ,N'{thanhpho}'
#                         ,N'{dantoc}'
#                         ,N'Việt Nam'
#                         ,N'{tongiao}'
#                         ,N'{trinhdo}'
#                         ,N'{quequan}'
#                         ,N'{tamtru}'
#                         ,NULL
#                         ,N'{masothue}'
#                         ,N'{nganhang}'
#                         ,N'{sotaikhoan}'
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,'{sodienthoainguoithan}'
#                         ,N'{hopdong}'
#                         ,'{ngaychinhthuc}'
#                         ,NULL
#                         ,N'{jobtitlevn}'
#                         ,N'{hccategory}'
#                         ,N'{gradecode}'
#                         ,N'{factory}'
#                         ,N'{department}'
#                         ,N'{chucvu}'
#                         ,N'{sectioncode}'
#                         ,N'{section}'
#                         ,N'{line}'
#                         ,N'{employeetype}'
#                         ,N'{jobtitleen}'
#                         ,N'{positioncode}'
#                         ,N'{positioncodedescription}'
#                         ,'5000000'
#                         ,'Không'
#                         ,NULL
#                         ,'{ngaythuviec}'
#                         ,NULL
#                         ,N'{tinhtranglamviec}'
#                         ,NULL
#                         ,'1'
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                         ,NULL
#                     )
#                 """
#                 # print(query)
#                 try:
#                     print(cursor.execute(query))
#                     conn.commit()
#                 except Exception as e:
#                     print(query)
#                     print(e)
#             conn.close()
    
put_nt1()