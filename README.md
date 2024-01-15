# WeatherXL

Thu thập dữ liệu lịch sử thời tiết hàng ngày trực tuyến từ nguồn freeMeteo và TimeAndDate

[Nhấn tải WeatherXL](https://github.com/SanbiVN/WeatherXL/releases/download/weather/WeatherXL_v3.04.zip)  

[![Tổng tải xuống](https://img.shields.io/github/downloads/SanbiVN/WeatherXL/total.svg)](https://github.com/SanbiVN/WeatherXL/releases/download/weather/WeatherXL_v3.04.zip)


![WeatherXL](https://github.com/SanbiVN/WeatherXL/blob/main/images/meteo%20weather.gif)


# HƯỚNG DẪN SỬ DỤNG
​
Để Add-in lấy được thông tin từ Trang tính để thực hiện cập nhật dữ liệu, cần thực hiện các bước như hướng dẫn dưới đây.​
​
Tạo các ô với các Name như sau: (Trong tab Formulas > Name Manager)​
(Tạo name để tự động lấy thông tin tải và ghi dữ liệu)​
​
### Các ô bắt buộc:​
1. *Nguồn web*: tên ```tt_Nguon​```
2. *Từ ngày*: tên ```tt_TuNgay​```
3. *Đến ngày*: tên ```tt_DenNgay​```
(Không cần tạo name ```tt_TuNgay``` và ```tt_DenNgay``` khi có cột dữ liệu tên ```tt_TheoNgay```)​
​
​
### Các ô cột dữ liệu:​
Đặt ô với Name tên tt_DuLieu đại diện vùng sẽ ghi toàn bộ dữ liệu thông tin thời tiết vào trang tính.​ \
Nếu muốn dữ liều các cột riêng lẻ hãy tạo Name như dưới đây, các Name phải tạo tại tiêu đề cột, Name nào không có thì bỏ qua không ghi dữ liệu:

4. *Nhiệt độ (nhỏ - lớn)*: tên ```tt_NhietDo``` (giá trị: 24 / 28)​
5. *Nhiệt độ (<) nhỏ nhất*: tên ```tt_NhietDo_Nho​```
6. *Nhiệt độ (<) lớn nhất*: tên ```tt_NhietDo_Lon​```
7. *Ngày hoặc Theo Ngày*: ​
+ Nếu tên ```tt_Ngay_Tang```: với cột ngày sắp xếp tăng dần​
+ Nếu tên ```tt_Ngay_Giam```: với cột ngày sắp xếp giảm dần​
+ Nếu tên ```tt_TheoNgay```: khi dữ liệu cột ngày đã tồn tại, dữ liệu sẽ ghi vào dòng tương ứng ngày​
(Không cần tạo name ```tt_TuNgay``` và ```tt_DenNgay```)​
8. *Mức gió ổn định tối đa*: tên ```tt_MucGio​```
9. *Gió giật tối đa*: tên ```tt_GioGiat​```
10. *Lượng mưa*: tên ```tt_LuongMua​```
11. *Mô tả*: tên ```tt_MoTa​```
12. *Hiện icon*: tên ```tt_icon``` (Nếu name tồn tại thì Icon sẽ được thêm vào dòng dữ liệu)​
​
Để tạo tất cả Name trên nhanh hơn, hãy gõ hàm ```=ThoiTiet_AddNames()``` vào ô bất kì, các Name sẽ tự động được tạo và cửa sổ ```Name Manager``` sẽ hiện lên để chỉnh sửa.​
​
Để tạo trang tính có sẵn, hãy gõ hàm ```=ThoiTiet_Worksheet()``` vào ô bất kì,​
Một trang tính mới sẽ được tạo vào dự án của bạn với giao diện đầy đủ.​
​
​
Tìm vị trí vùng địa lý:​
​
Gõ hàm ```=ThoiTiet_TimKiem("Hà Nội")``` sẽ tìm kiếm *vùng+đường dẫn* và ghi vào tại vị trí ô gõ hàm.​
Bạn cần chép đường dẫn vị trí địa lý cần thiết vào ô Name ```tt_Nguon```.​


Gán nút nhấn:

Chép mã dưới đây vào module dự án của bạn, để gán nút cập nhật dữ liệu.​
Mã sẽ tự động tìm kiếm Add-in và thực thi các lệnh.​

```VBA
        Sub GetWeatherVN(Optional Direction&)
              WeatherXLCall "GetWeatherVN", Direction
        End Sub
        Sub ClearWeatherVN(Optional Direction&)
              WeatherXLCall "ClearWeatherVN", Direction
        End Sub
        Sub sortDataMeteoWeather(Optional Direction&)
              WeatherXLCall "sortDataMeteoWeather", Direction
        End Sub
        Sub sortDataTADWeather(Optional Direction&)
              WeatherXLCall "sortDataTADWeather", Direction
        End Sub
        Private Sub WeatherXLCall(Byval proc$, Optional Direction&)
            On Error Resume Next
            Dim a
            For Each a In Application.AddIns
                If a.Name Like "WeatherXL*" Then
                    Application.OnTime Now, "'" & a.Name & "'!'" & proc & " " & Direction & "'": Exit Sub
                End If
            Next
            MsgBox "Hay cai dat Add-in WeatherXL", vbInformation
            Err.clear
        End Sub
```
Gán tên GetWeatherVN vào nút nhấn cập nhật dữ liệu


### Phiên bản cập nhật:
Trình tự động tìm kiếm bản cập nhật mới nhất tại Github​ 
- Để tìm bản cập nhật mới gõ hàm: ```=ThoiTiet_Update()​``` 
- Để tắt gõ hàm: ```=ThoiTiet_UpdateOff()​``` 
- Để bật gõ hàm: ```=ThoiTiet_UpdateOn()```
