#encoding:utf-8
#author: anion

module ReportConfiguration
  class ExcelWin
    @@id=9
    @@name=''
    def initialize
      @excel
    end
    def excel_new(encoding='utf-8')
      initialize
      @@worksheets_name =[]
      @excel = WIN32OLE.new("EXCEL.APPLICATION")
      @excel.Visible=true
      @workbook = @excel.WorkBooks.Add()
      @encoding = encoding
    end

    def excel_sheet_name(name)
      while @@worksheets_name.include?(name)
        name +="1"
      end
      @@worksheets_name << name
      worksheet = @workbook.Worksheets.Add()
      worksheet.Activate
      worksheet.name = name
    end

    def create_report_file(date,start_time,end_time)
      excel_new
      @excel.DisplayAlerts = false

      @objSheet =  @excel.Sheets.Item(1)
      @excel.Sheets.Item(1).Select
      @objSheet.Name = "自动化测试报告"

      @objSheet.Range("B1").Value = "测试结果"
      #合并单元格
      @objSheet.Range("B1:E1").Merge
      #水平居中 -4108
      @objSheet.Range("B1:E1").HorizontalAlignment = -4108
      @objSheet.Range("B1:E1").Interior.ColorIndex = 53
      @objSheet.Range("B1:E1").Font.ColorIndex = 5
      @objSheet.Range("B1:E1").Font.Bold = true
      @objSheet.Range("B1:E1").Font.Size =24

      # @objSheet.Range("B2:E2").Merge
      # @objSheet.Rows(2).RowHeight = 20

      rowNum = [2,3,4,5,6,7]
      rowNum.each {|re|
        @objSheet.Range("C#{re}:E#{re}").Merge}

      # @objSheet.Range("B9:E9").Merge
      # @objSheet.Rows(9).RowHeight = 30

      #Set the Date and time of Execution
      @objSheet.Range("B2").Value = "测试日期: "
      @objSheet.Range("B3").Value = "开始时间: "
      @objSheet.Range("B4").Value = "结束时间: "
      @objSheet.Range("B5").Value = "持续时间: "
      @objSheet.Range("C2").Value = date
      @objSheet.Range("C3").Value = start_time
      @objSheet.Range("C4").Value = end_time
      @objSheet.Range("C5").Value = "=R[-1]C-R[-2]C"
      # @objSheet.Range("C5").NumberFormat ="[h]:mm:ss;@"

      #Set the Borders for the Date & Time Cells
      @objSheet.Range("B2:E7").Borders(1).LineStyle = 1
      @objSheet.Range("B2:E7").Borders(2).LineStyle = 1
      @objSheet.Range("B2:E7").Borders(3).LineStyle = 1
      @objSheet.Range("B2:E7").Borders(4).LineStyle = 1

      #Format the Date and Time Cells
      @objSheet.Range("B2:E7").Interior.ColorIndex = 40
      @objSheet.Range("B2:E7").Font.ColorIndex = 14
      @objSheet.Range("B2:E7").Font.Bold = true

      #Track the Row Count and insrtuct the viewer not to disturb this
      @objSheet.Range("C6").AddComment
      @objSheet.Range("C6").Comment.Visible = false
      @objSheet.Range("C6").Comment.Text "这点生成的数据大家不要删除哦"
      @objSheet.Range("C6").Value = "0"
      @objSheet.Range("B6").Value = "自动化测试场景总数:"
      @objSheet.Range("B7").Value = "测试通过率:"

      @objSheet.Range("B8").Value = "自动化RB文件名"
      @objSheet.Range("C8").Value = "测试场景ID"
      @objSheet.Range("D8").Value = "测试场景功能"
      @objSheet.Range("E8").Value = "执行结果"

      #添加超链接功能
      # @objSheet.Hyperlinks.Add(@objSheet.Range("B9"), "","测试结果!A1")
      # @objSheet.Range("B9").Value = "点击测试用例名称打开详情页面."

      #  @objSheet.Hyperlinks.Add(@objSheet.Range("B9"), "http://www.163.com")
      #Format the Heading for the Result Summery
      @objSheet.Range("B8:E8").Interior.ColorIndex = 53
      @objSheet.Range("B8:E8").Font.ColorIndex = 1
      @objSheet.Range("B8:E8").Font.Bold = true
      @objSheet.Range("B8:E8").Font.Size = 13

      #Set the Borders for the Result Summery
      @objSheet.Range("B8:E8").Borders(1).LineStyle = 1
      @objSheet.Range("B8:E8").Borders(2).LineStyle = 1
      @objSheet.Range("B8:E8").Borders(3).LineStyle = 1
      @objSheet.Range("B8:E8").Borders(4).LineStyle = 1

      #Set Column width
      @objSheet.Columns("B:E").Select

      @objSheet.Range("B9").Select
      @objSheet.Range("B9").ColumnWidth=27
      @objSheet.Range("C9").ColumnWidth=12
      @objSheet.Range("D9").ColumnWidth=55
      @objSheet.Range("E9").ColumnWidth=12

      # #Freez pane
      @excel.ActiveWindow.FreezePanes = true
    end

    def fill_in_auto_data(name,desc,res)
      unless @@name==name
        @objSheet.Range("B#{@@id}").Value=name
      end
      @@name=name
      @objSheet.Range("C#{@@id}").Value=@@id-8
      @objSheet.Range("D#{@@id}").Value=desc
      @objSheet.Range("E#{@@id}").Value=res
      if res=='passed'
        @objSheet.Range("E#{@@id}").Interior.ColorIndex = 4
      else
        @objSheet.Range("E#{@@id}").Interior.ColorIndex = 3
      end
      @@id +=1
      @@id
    end

    def merge_sheet(st,ed)
      @objSheet.Range("B#{st}:B#{ed}").Merge

    end

    def case_sheet_style(ed,sum,percent)
      @objSheet.Range("C6").Value=sum
      @objSheet.Range("C7").Value=percent
      @objSheet.Range("B9:E#{ed}").Borders(1).LineStyle = 1
      @objSheet.Range("B9:E#{ed}").Borders(2).LineStyle = 1
      @objSheet.Range("B9:E#{ed}").Borders(3).LineStyle = 1
      @objSheet.Range("B9:E#{ed}").Borders(4).LineStyle = 1
      @objSheet.Range("B9:E#{ed}").Font.Size = 12
    end


    def excel_save(file_path)
      #Save the Workbook at the specified Path with the Specified Name
      if File.exist?(file_path)
        File.delete(file_path)
      end
      @excel.ActiveWorkbook.SaveAs(file_path)
      @excel.Quit()
    end

    def excel_quit
      # @excel.Quit                      # 退出当前Excel文件
      @workbook.close(1)                       #关闭表sheet空间
      # exec('taskkill /f /im Excel.exe ')   #强制关闭所有的Excel进程
    end
  end
end