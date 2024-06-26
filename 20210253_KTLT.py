#Họ tên: Quách Thái Dương
#MSSV: 20210253

import pandas as pd 

##__________________MENU sử dụng trong chương trình______________________##

#Menu chính của chương trình
def menu():
    print("")
    print(" " * 20, "_" * 53)
    print(" " * 20, "|{0:51}|".format(" "))
    print(" " * 20, "|{0:^9}{1:^9}{0:^15}|".format(" ", "1. Nhập thông tin sinh viên"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^15}|".format(" ", "2. Xuất danh sách sinh viên"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^13}|".format(" ", "3. Sắp xếp sinh viên theo lớp"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^11}|".format(" ", "4. Cập nhật thông tin sinh viên"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^5}|".format(" ", "5. Tìm sinh viên theo mã số sinh viên"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^17}|".format(" ", "6. Tìm sinh viên theo lớp"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^14}|".format(" ", "7. Tìm sinh viên theo họ tên"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^16}|".format(" ", "8. Xóa sinh viên theo MSSV"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^9}|".format(" ", "9. Thoát khỏi chương trình và lưu"))
    print(" " * 20, "|{0:51}|".format(" "))
    print(" " * 20, "-" * 53)
    print("")
    print("")

#Menu cho tính năng cập nhật thông tin sinh viên
def menuCapNhat():
    print("")
    print(" " * 20, "_" * 64)
    print(" " * 20, "|{0:62}|".format(" "))
    print(" " * 20, "|{0:^9}{1:^9}{0:^35}|".format(" ", "1. Mã số sinh viên"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^44}|".format(" ", "2. Họ tên"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^8}|".format(" ", "3. Ngày sinh (Nhập theo định dạng dd/mm/yyyy)"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^41}|".format(" ", "4. Ngành học"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^41}|".format(" ", "5. Khoa viện"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^9}|".format(" ", "6. Lớp (Nhập theo định dạng, VD: K66 MI1-02)"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^37}|".format(" ", "7. Dừng cập nhật"))
    print(" " * 20, "|{0:62}|".format(" "))
    print(" " * 20, "-" * 64)
    print("")
    print("")

#Menu cho tính năng nhập mới thông tin sinh viên
def menuNhapTTin():
    print("")
    print(" " * 20, "_" * 47)
    print(" " * 20, "|{0:45}|".format(" "))
    print(" " * 20, "|{0:^9}{1:^9}{0:^4}|".format(" ", "1. Nhập thông tin từng sinh viên"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^14}|".format(" ", "2. Lấy dữ liệu từ file"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^25}|".format(" ", "3. Quay lại"))
    print(" " * 20, "|{0:45}|".format(" "))
    print(" " * 20, "-" * 47)
    print("")
    print("")

#Menu cho tính năng lưu và thoát khỏi chương trình
def menuSaveExit():
    print("")
    print(" " * 20, "_" * 28)
    print(" " * 20, "|{0:26}|".format(" "))
    print(" " * 20, "|{0:^9}{1:^4}{0:^12}|".format(" ", "1. Có"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^8}|".format(" ", "2. Không"))
    print(" " * 20, "|{0:^9}{1:^9}{0:^6}|".format(" ", "3. Quay lại"))
    print(" " * 20, "|{0:26}|".format(" "))
    print(" " * 20, "-" * 28)
    print("")
    print("")

#Thuộc tính cho dữ liệu của sinh viên 
class SinhVien:
    def __init__(self, mssv, hoTen, ngaySinh, nganhHoc, khoaVien, lopHoc):
        self.mssv = mssv
        self.hoTen = hoTen
        self.ngaySinh = ngaySinh
        self.nganhHoc = nganhHoc
        self.khoaVien = khoaVien
        self.lopHoc = lopHoc
        self.next = None

   
class QLSV:
    def __init__(self):
        self.head = None

#kiểm tra dữ liệu nhập vào có là số không
#Đầu vào: Giá trị bât kỳ
#Đầu ra: True nếu đúng là số, False ngược lại

    def kiemTraSo(self, value):
        try:
            int(value)
            return True
        except ValueError:
            return False    


#Hàm hỗ trợ đưa họ tên khi nhập vào về đúng dạng tên thông thường
#Đầu vào: Xâu 
#Đầu ra: Xâu về đúng dạng tên 

    def normalName(self, string):
        string = string.split()  #chuyển xâu ban đầu thành một list có phần tử là các từ
        result = ""     #biến trả về xâu chuẩn hóa tên thông thường
        for i in range(len(string)):
            string[i] = string[i].capitalize() #viết hoa kí tự đầu mỗi từ
            result += string[i]
            result += " "
        result = result.strip() #loại bỏ khoảng trắng 2 đầu
        return result
    
#Hàm hỗ trợ đưa ngày tháng về đúng định dạng chương trình yêu cầu
#Đầu vào: Ngày 
#Đầu ra: Ngày đã được hiệu chỉnh
    def normalDate(self, string):
        date = ""  #Biến hiệu chỉnh ngày sinh nhập vào
        for i in string: #i: Biến chạy xâu newData
            if i == "-":
                date += "/"
            elif i == " ":
                date += ""
            else:
                date += i
        return date
    
#Hàm hỗ trợ kiểm tra định dạng tên lớp nhập vào
#Đầu vào: Tên lớp
#Đầu ra: True hoặc False
    def kiemTraLop(self, string):
        listLop = string.split()
        if len(listLop) == 2 or len(listLop) == 0:
            return True
        else:
            return False

##___________________Nhập thông tin sinh viên__________________## 
#Nhập dữ liệu vào danh sách liên kết đơn
#Đầu vào: Thông tin của 1 sinh viên

    def nhapDuLieuSV(self, mssv, hoTen, ngaySinh, nganhHoc, khoaVien, lopHoc):
        duLieuSinhVien = SinhVien(mssv, hoTen, ngaySinh, nganhHoc, khoaVien, lopHoc)
        if self.head == None:
            self.head = duLieuSinhVien
        else:
            current = self.head  #current: biến chạy duyệt qua danh sách
            while current.next:
                current = current.next
            current.next = duLieuSinhVien

#Nhập thông tin từ file vào danh sách liên kết đơn
#Đầu vào: Đường dẫn tới file truy cập

    def nhapDuLieuTuFile(self, path):
        data = pd.read_excel(path) #data: biến lưu dữ liệu sau khi đọc file
        for _, r in data.iterrows():
            mssv = int(r["MSSV"])
            hoTen = r["Họ tên"]
            ngaySinh = r["Ngày sinh"]
            nganhHoc = r["Ngành"]
            khoaVien = r["Khoa viện"]
            lopHoc = r["Lớp"]
            self.nhapDuLieuSV(mssv, hoTen, ngaySinh, nganhHoc, khoaVien, lopHoc)

#Nhập thông tin và lọc từ file được chọn vào dữ liệu lưu
#Đầu vào: Đường dẫn đến file truy cập
#Đầu ra: Thông báo thêm file hoặc link không hợp lệ
#Lưu ý: Chương trình không thêm các sinh viên có mssv đã có trong file và chỉ lấy dữ liệu xuất hiện đầu trong 2 dữ liệu trùng từ file

    def nhapThongTinTuFile(self, path):
        try:
            data = pd.read_excel(path) #data: biến lưu dữ liệu sau khi đọc file
        except FileNotFoundError:
            print("Link không hợp lệ")
        except OSError:
            print("Link không hợp lệ")
        else:
            for _,r in data.iterrows():
                mssv = int(r["MSSV"])
                hoTen = r["Họ tên"]
                ngaySinh = r["Ngày sinh"]
                nganhHoc = r["Ngành"]
                khoaVien = r["Khoa viện"]
                lopHoc = r["Lớp"]
                current = self.head 

                # flag: gắn cờ báo hiệu có sinh viên trùng mssv  
                flag = True  
                while current:  ##current: biến chạy duyệt qua danh sách
                    if mssv == current.mssv:
                        flag = False
                    current = current.next
                if not self.kiemTraLop(lopHoc):
                    flag = False
                if flag == True:
                    hoTen = self.normalName(hoTen)
                    ngaySinh = self.normalDate(ngaySinh)
                    lopHoc = self.normalName(lopHoc)
                    lopHoc = lopHoc.upper()
                    self.nhapDuLieuSV(mssv, hoTen, ngaySinh, nganhHoc, khoaVien, lopHoc)
            print("Thông tin từ file đã được thêm, nhấn 9 hoàn tất cập nhật")
            

#Nhập thông tin từng sinh viên 
#Đầu ra: Thông tin được nhập vào chương trình

    def nhapThongTinSV(self):
        while True:
            soLuong = input("Nhập số lượng sinh viên cần nhập thông tin: ")
            if not self.kiemTraSo(soLuong):
                print("Nhập số hợp lệ!")
                continue
            soLuong = int(soLuong)
            for i in range(soLuong): #i: biến đếm lần nhập thông tin sinh viên
                print("Nhập thông tin sinh viên thứ", i+1)
                while True:
                    mssv = input("Nhập mã số sinh viên: ")

                    # current: biến chạy trong danh sách liên kết đơn 
                    current = self.head
                    flag = True

                    # flag: cờ báo hiệu đã có mssv trong danh sách
                    if not self.kiemTraSo(mssv):
                        print("Nhập lại mã số sinh viên")
                        continue
                    while current:
                        mssv = int(mssv)
                        if current.mssv == mssv :
                            flag = False
                        current = current.next
                    if flag == False:
                        print("Mã số sinh viên đã có trong danh sách hãy nhập lại")  
                        continue
                    break
                hoTen = input("Nhập họ tên: ")
                if hoTen != "":
                    hoTen = self.normalName(hoTen)
                while True:
                    nam = input("Nhập năm sinh: ")
                    if nam == "":
                        break
                    if not self.kiemTraSo(nam):
                        print("Nhập lại năm sinh")
                        continue
                    break 
                while True:
                    thang = input("Nhập tháng sinh: ")
                    if thang == "":
                        break
                    if not self.kiemTraSo(thang):
                        print("Nhập lại tháng sinh")
                        continue
                    thangInt = int(thang)
                    if thangInt > 12 or thangInt < 1:
                        print("Tháng sinh không hợp lệ, hãy nhập lại")
                        continue
                    break
                while True:
                    ngay = input("Nhập ngày sinh: ")
                    if ngay == "":
                        break
                    if not self.kiemTraSo(ngay):
                        print("Nhập lại ngày sinh")
                        continue
                    ngayInt = int(ngay)
                    if ngayInt > 31 or ngayInt < 1:
                        print("Ngày sinh không hợp lệ, hãy nhập lại")
                        continue
                    break 
                ngay = ngay.strip()
                thang = thang.strip()
                nam = nam.strip()
                ngaySinh = ngay + "/" + thang + "/" + nam
                nganhHoc = input("Nhập tên ngành học: ")
                khoaVien = input("Nhập tên khoa viện đang học: ")
                while True:
                    lopHoc = input("Nhập tên lớp (Ví dụ: K66 MI1-02): ")
                    if not self.kiemTraLop(lopHoc):
                        print("Nhập lại tên lớp đúng định dạng")
                        continue
                    break   
                lopHoc = self.normalName(lopHoc)
                lopHoc = lopHoc.upper()
                self.nhapDuLieuSV(mssv, hoTen, ngaySinh, nganhHoc, khoaVien, lopHoc)
                print("Thông tin sinh viên đã được thêm, chọn 9 để hoàn tất cập nhật")
            break
##__________Xuất danh sách sinh viên___________##
#Đầu ra: Danh sách thông tin sinh viên

    def xuatDanhSachSV(self):
        
        if self.head is None:
            print("Không có danh sách sinh viên")
        
        else:
            current = self.head #current: biến chạy duyệt qua danh sách

            #records: list lưu trữ thông tin sinh viên 
            records = []
            while current:
                record = {
                    "MSSV": current.mssv,
                    "Họ tên": current.hoTen,
                    "Ngày sinh": current.ngaySinh,
                    "Ngành": current.nganhHoc,
                    "Khoa viện": current.khoaVien,
                    "Lớp": current.lopHoc
                }
                records.append(record)
                current = current.next
            
            print("Danh sách sinh viên: ")
            #df: biến in ra bảng dữ liệu sinh viên
            df = pd.DataFrame(records)
            print(df)
    
#_______Sắp xếp danh sách sinh viên____________#
#Đầu vào: Danh sách sinh viên
#Đẩu ra: Danh sách sinh viên sắp xếp theo lớp

#hàm trả về phần tử giữa danh sách liên kết đơn
#mid: phần tử ở giữa
    def findMid(self, start):
        if self.head is None:
            return start
        mid = start
        end = start
        while end.next and end.next.next:
            mid = mid.next
            end = end.next.next
        
        return mid

#hàm trả về danh sách đã được gộp lại    
    def merge(self, left, right):
        result = None

        if left is None:
            return right
        if right is None:
            return left
        
        if left.lopHoc <= right.lopHoc:
            result = left
            result.next = self.merge(left.next, right)
        else:
            result = right
            result.next = self.merge(left, right.next)    

        return result
    
    def mergeSort(self, start):
        if start is None or start.next is None:
            return start
        mid = self.findMid(start)
        nextMid = mid.next
        mid.next = None

        left = self.mergeSort(start)
        right = self.mergeSort(nextMid)

        sortList = self.merge(left,right)
        return sortList

#hàm lưu danh sách đã sắp xếp vào self ban đầu   
    def disMergeSort(self):
        self.head = self.mergeSort(self.head)

#hàm in danh sách trước và sau khi sắp xếp
    def sapXepSV(self):
        if self.head:
            print("Trước khi sắp xếp:")
            self.xuatDanhSachSV()

            self.disMergeSort()

            print("Sau khi sắp xếp:")
            self.xuatDanhSachSV()
        else:
            print("Không có danh sách để sắp xếp!")

# flag trong các hàm tìm kiếm cập nhật là cờ báo hiệu có sinh viên cần tìm

#_______Tìm kiếm sinh viên theo mã số sinh viên_______#
#Đầu vào: Mã số sinh viên
#Đầu ra: Thông tin sinh viên hoặc báo không tìm thấy

    def timMSSV(self):
        if self.head:
            mssvCanTim = input("Nhập mã số sinh viên cần tìm: ")
            if not self.kiemTraSo(mssvCanTim):
                print("Nhập đúng định dạng mã số sinh viên")
            else:
                mssvCanTim = int(mssvCanTim)
                flag = False
                records = []
                current = self.head  #current: biến chạy duyệt qua danh sách
                while current:
                    if current.mssv == mssvCanTim:
                        record = {"MSSV": current.mssv,
                                "Họ tên": current.hoTen,
                                "Ngày sinh": current.ngaySinh,
                                "Ngành": current.nganhHoc,
                                "Khoa viện": current.khoaVien,
                                "Lớp": current.lopHoc}
                        records.append(record)
                        flag = True
                        break
                    current = current.next
                if flag == True:
                    print("Thông tin sinh viên có mssv thỏa mãn: ")
                    df = pd.DataFrame(records) #df: biến in ra bảng dữ liệu sinh viên
                    print(df)
                else:
                    print("Không tìm thấy thông tin sinh viên với mssv này")
        else:
            print("Không có danh sách sinh viên để tìm kiếm")

#_________Hàm hỗ trợ tìm kiếm sinh viên theo lớp và họ tên__________#
#Đầu vào: Xâu
#Đầu ra: Xâu đã xóa khoảng trắng, kí tự "-" và viết thường

    def supportSearch(self, string):
        result = ""         #result: biến lưu kết quả để trả về
        for kiTu in string: #kiTu: kí tự xâu
            if kiTu == " " or kiTu == "-":
                result += ""
            else:
                result += kiTu
        result = result.lower()
        return result
    
#______Tìm kiếm sinh viên theo lớp________#
#Đầu vào: Lớp sinh viên
#Đầu ra: Thông tin sinh viên hoặc báo không tìm thấy

    def timLop(self):
        if self.head:
            lopHocCanTim = input("Nhập tên lớp học (Đúng định dạng ví dụ: K66 MI1-02): ")
            string1 = self.supportSearch(lopHocCanTim) #string1: xâu đại diện cho lopHocCanTim để so sánh
            flag = False
            records = []
            current = self.head
            while current:          #current: biến chạy duyệt qua danh sách
                string2 = self.supportSearch(current.lopHoc) #string2: xâu đại diện cho current.lopHoc để so sánh
                if string1 == string2:
                    record = {"MSSV": current.mssv,
                            "Họ tên": current.hoTen,
                            "Ngày sinh": current.ngaySinh,
                            "Ngành": current.nganhHoc,
                            "Khoa viện": current.khoaVien,
                            "Lớp": current.lopHoc}
                    records.append(record)
                    flag = True
                current = current.next
            if flag == True:
                print("Thông tin sinh viên có lớp học thỏa mãn: ")
                df = pd.DataFrame(records) #df: biến in ra bảng dữ liệu sinh viên
                print(df)
            else:
                print("Không tìm thấy thông tin sinh viên với lớp học này")
        else:
            print("Không có danh sách sinh viên để tìm kiếm")


#_______Tìm kiếm sinh viên theo họ tên______#
#Đầu vào: Họ tên sinh viên
#Đầu ra: Thông tin sinh viên hoặc báo không tìm thấy

    def timTen(self):
        if self.head:
            hoTenCanTim = input("Nhập họ tên: ")
            string1 = self.normalName(hoTenCanTim) #string1: Xâu hiệu chỉnh của xâu hoTenCanTim để so sánh
            flag = False
            records = []
            current = self.head
            while current:        
                if string1 == current.hoTen:
                    record = {"MSSV": current.mssv,
                              "Họ tên": current.hoTen,
                              "Ngày sinh": current.ngaySinh,
                              "Ngành": current.nganhHoc,
                              "Khoa viện": current.khoaVien,
                              "Lớp": current.lopHoc}
                    records.append(record)
                    flag = True
                current = current.next
            if flag == True:
                print("Thông tin sinh viên có họ tên thỏa mãn: ")
                df = pd.DataFrame(records) #df: biến in ra bảng dữ liệu sinh viên
                print(df)
            else:
                print("Không tìm thấy thông tin sinh viên với họ tên này")
        else:
            print("Không có danh sách sinh viên để tìm kiếm")
        


#_______Cập nhật thông tin sinh viên____________#
#Đầu vào: Mã số sinh viên cần cập nhật, Thông tin mới
#Đầu ra: Cập nhật thông tin sinh viên

    def capNhat(self):
        if self.head is not None:
            self.xuatDanhSachSV()
            mssvCapNhat = input("Nhập mã số sinh viên của sinh viên cần cập nhật: ")
            flag = False
            if not self.kiemTraSo(mssvCapNhat):
                print("Nhập đúng định dạng số cho mã số sinh viên")
            else:
                mssvCapNhat = int(mssvCapNhat)
                current = self.head         #current: biến chạy duyệt qua danh sách
                while current: 
                    if current.mssv == mssvCapNhat:
                        flag = True
                        menuCapNhat()                        
                        while True: 
                            choose = input("Lựa chọn số tương ứng thông tin cần cập nhật:")     #choose: biến lựa chọn tính năng
                            
                            if not self.kiemTraSo(choose):
                                print("Thao tác không hợp lệ")
                            else:
                                choose = int(choose)
                                if choose == 7:
                                    print("Hoàn thành cập nhật thông tin, nhấn 9 để lưu vào file")
                                    break

                                elif choose < 1 or choose > 7:
                                    print("Thao tác không hợp lệ")

                                else:
                                    newData = input("Nhập thông tin mới: ") #newData: biến lưu giá trị thông tin mới
                                    if choose == 1:
                                        if not self.kiemTraSo(newData):
                                            print("Thông tin nhập vào không hợp lệ")
                                        else:
                                            newData = int(newData)
                                        # flag2: cờ báo hiệu mã số sinh viên xuất hiện trong danh sách
                                            flag2 = True
                                            current1 = self.head
                                            while current1:
                                                if current1.mssv == newData and current.mssv != newData:
                                                    print("Mã số sinh viên này đã xuất hiện trong danh sách không thể cập nhật")
                                                    flag2 = False
                                                current1 = current1.next
                                            if flag2 == True:
                                                current.mssv = newData
                                
                                    elif choose == 2:
                                        newData = self.normalName(newData)
                                        current.hoTen = newData
                                    
                                    elif choose == 3:
                                        newData = self.normalDate(newData)
                                        current.ngaySinh = newData
                                    
                                    elif choose == 4:
                                        current.nganhHoc = newData
                                    
                                    elif choose == 5:
                                        current.khoaVien = newData

                                    elif choose == 6:
                                        if not self.kiemTraLop(newData):
                                            print("Nhập sai định dạng lớp")
                                        else:
                                            newData = self.normalName(newData)
                                            newData = newData.upper()
                                            current.lopHoc = newData
                    current = current.next
                        
                if flag != True:
                    print("Không tìm thấy sinh viên cần cập nhật thông tin")

        else:
            print("Không có danh sách sinh viên để tìm kiếm")

#_______Xóa sinh viên theo mã số sinh viên_______#
#Đầu vào: Mã số sinh viên của sinh viên cần xóa
#Đầu ra: Danh sách sinh viên sau khi xóa

    def xoaSV(self):
        if self.head is None:
            print("Không có danh sách sinh viên để thực hiện xóa")
        else:
            self.xuatDanhSachSV()
            mssvXoa = input("Nhập MSSV của sinh viên cần xóa: ")
            if not self.kiemTraSo(mssvXoa):
                print("Nhập đúng định dạng số cho mã số sinh viên")
            else:  
                mssvXoa = int(mssvXoa)
                current = self.head
                
                #Xóa phần tử ở đầu danh sách
                if current and current.mssv == mssvXoa:
                    self.head = current.next
                    current = None
                    print("Danh sách sau khi xóa")
                    self.xuatDanhSachSV()
                    print("Nhấn 9 để lưu thao tác xóa sinh viên trong file")
                    print("Nhấn 8 để xóa tiếp")
                #Xóa phần tử không ở đầu
                else:
                    while current:
                        if current.mssv == mssvXoa:
                            break
                        prev = current
                        current = current.next
                    if current is None:
                        print("Không tìm thấy sinh viên cần xóa")
                    else:
                        prev.next = current.next
                        current = None
                        print("Danh sách sau khi xóa")
                        self.xuatDanhSachSV()
                        print("Nhấn 9 để lưu thao tác xóa sinh viên trong file")
                        print("Nhấn 8 để xóa tiếp")

#_____________Lưu các thao tác đã thực hiện vào file__________#            
#Đầu vào: Đường dẫn tới file để lưu
#Đầu ra: Thông tin được lưu vào file

    def luuThaoTacVaoFile(self, path):
        records = []
        current = self.head
        while current:
            record = {
                "MSSV": current.mssv,
                "Họ tên": current.hoTen,
                "Ngày sinh": current.ngaySinh,
                "Ngành": current.nganhHoc,
                "Khoa viện": current.khoaVien,
                "Lớp": current.lopHoc
            }
            records.append(record)
            current = current.next
        
        df = pd.DataFrame(records) #df: biến in ra bảng dữ liệu sinh viên
        df.to_excel(path, index = False) 
        print("Thao tác đã được lưu vào file")

#________________Truy cập chương trình________________#

if __name__ == "__main__":
    
    # path: biến lưu đường dẫn tới file
    path = input("Nhập đường dẫn tới file cần làm việc (file: dulieusv.xlsx): ")
    qlsv =QLSV()
    try:
        qlsv.nhapDuLieuTuFile(path)
    except FileNotFoundError:
        print("Link không hợp lệ")
    except OSError:
        print("Link không hợp lệ")
    else:
        #flag: cờ giúp kết thúc vòng lặp khi chuyển thành False
        flag = True
        while flag:
            menu()
            choose = input("Nhập số để lựa chọn tính năng: ")
            if not qlsv.kiemTraSo(choose):
                print("Nhập lại số tương ứng")
                continue
            choose = int(choose)
            if choose == 1:
                print("Bạn muốn nhập thông tin từng sinh viên hay lấy dữ liệu từ file khác?")
                
                while True:
                    menuNhapTTin()
                    choose3 = input("Nhập số lựa chọn:")
                    if not qlsv.kiemTraSo(choose3):
                        print("Nhập lại số tương ứng")
                        continue
                    choose3 = int(choose3)
                    if choose3 == 1:
                        qlsv.nhapThongTinSV()
                        break
                    elif choose3 == 2:
                        # link: biến lưu đường link được nhập
                        link = input("Nhập đường dẫn tới file dulieumoi.xlsx: ")
                        qlsv.nhapThongTinTuFile(link)  
                        break
                    elif choose3 == 3:
                        break
                    else:
                        print("Thao tác không hợp lệ")
                        continue

            elif choose == 2:
                qlsv.xuatDanhSachSV()

            elif choose == 3:
                qlsv.sapXepSV()

            elif choose == 4:
                qlsv.capNhat()

            elif choose == 5:
                qlsv.timMSSV()

            elif choose == 6:
                qlsv.timLop()

            elif choose == 7:
                qlsv.timTen()

            elif choose == 8:
                qlsv.xoaSV()

            elif choose == 9:
                print("Bạn có muốn lưu thao tác?")
                
                while True:
                    menuSaveExit()
                    choose2 = input("Nhập số lựa chọn: ")
                    if not qlsv.kiemTraSo(choose2):
                        print("Nhập lại số tương ứng")
                        continue
                    choose2 = int(choose2)
                    if choose2 == 1:
                        qlsv.luuThaoTacVaoFile(path)
                        flag =False
                        break
                    elif choose2 == 2:
                        print("Hủy thao tác")
                        flag = False
                        break
                    elif choose2 == 3:
                        break
                    else:
                        print("Thao tác không hợp lệ")      
                                  
            else: 
                print("Thao tác không hợp lệ")
