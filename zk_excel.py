from zk import ZK, const
import xlsxwriter

conn = None
zk = ZK('192.168.1.10', port=4370, timeout=5)
try:
    print ('Connecting to device ...')
    conn = zk.connect()
    # Print Firmware version
    print(conn.get_firmware_version())

    # Get attendances (will return list of Attendance object)
    attendances = conn.get_attendance()

    # Create an Excel WorkBook
    workbook = xlsxwriter.Workbook('attendance_data.xlsx')
    worksheet = workbook.add_worksheet()

    # Controll worksheet row and columns
    row = 0
    col = 0

    # Add an Excel date format.
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd HH:mm:ss'})

    print('Writing to Excel File...')
    # Read Attendance Data and Write to Excel
    for attendance in attendances:
        worksheet.write(row, col, attendance.user_id)
        worksheet.write(row, col + 1, 'PYTHON_IMPORT')
        worksheet.write_datetime(row, col + 2, attendance.timestamp, date_format)
        row += 1
except:
    print ("Process terminate")
finally:
    workbook.close()
    if conn:
        conn.disconnect()
