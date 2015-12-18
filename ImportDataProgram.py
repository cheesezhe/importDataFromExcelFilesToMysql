#-*- coding: utf-8 -*-
import os,sys,datetime
import mysql.connector
import xlrd

'''
the main function to import data
    username: username of mysql database
    password: password for username
    database: a specific database in mysql
    datapath: the absolute path or relative path of data folder
'''
def importDataHelper(username, password, database, datapath):
    '''import data helper'''
    '''
    Step 0: Validate input database parameters
    '''
    try:
        conn = mysql.connector.connect(user=username, password=password, database=database, use_unicode=True)
    except mysql.connector.errors.ProgrammingError as e:
        print e
        return -1
    '''
    Step 1: Traverse files in datapath, store file paths and corresponding table names in lists
    lists[0] is the list of files paths
    lists[1] is the list of table names
    '''
    lists = getFilesList(datapath)
    nfiles = len(lists[0])
    '''
    Step 2: Store data in mysql via a for-loop
    '''
    cursor = conn.cursor()
    for file_idx in xrange(0, nfiles):
        file_path = lists[0][file_idx]
        print "processing file(%d/%d):[ %s ]"%(file_idx+1, nfiles, file_path)
        table_name = lists[1][file_idx]
        num = storeData(file_path, table_name, cursor)
        if num >= 0:
            print "[ %d ] data have been stored in TABLE:[ %s ]"%(num, table_name)
        conn.commit()
    cursor.close()
    '''
    Step 3: Close connection
    '''
    conn.close()

'''
get files list in the dir, including the files in its sub-folders
the return list contain two elements, the first element is a file names list
and the second element is a table names list(will be used for creating tables in database),
'''
def getFilesList(dir):
    path_list = []
    table_list = []
    file_name_list = os.listdir(dir)
    for file_name in file_name_list:
        path = os.path.join(dir, file_name)
        if os.path.isdir(path):
            '''get the files in sub folder recursively'''
            tmp_lists = getFilesList(path)
            path_list.extend(tmp_lists[0])
            table_list.extend(tmp_lists[1])
        else:
            path_list.append(path)
            '''convert file name to mysql table name'''
            file_name = file_name.split('.')[0] #remove .xls
            # file_name = file_name.split('from')[0] #remove characters after 'from'
            file_name = file_name.strip()#remove redundant space at both ends
            file_name = file_name.replace(' ','_') #replace ' ' with '_'
            file_name = file_name.replace('-','_') #replace ' ' with '_'
            file_name = file_name.lower() #convert all characters to lowercase
            table_list.append(file_name)
    return [path_list, table_list]

'''
store the data of file file_path in table table_name
    file_path: file location
    table_name: name of the table that will be created in database
    cursor: a mysql cursor
'''
def storeData(file_path, table_name, cursor):
    ret = 0
    '''open an excel file'''
    file = xlrd.open_workbook(file_path)
    '''get the first sheet'''
    sheet = file.sheet_by_index(0)
    '''get the number of rows and columns'''
    nrows = sheet.nrows
    ncols = sheet.ncols
    '''get column names'''
    col_names = []
    for i in range(0, ncols):
        title = sheet.cell(1, i).value
        title = title.strip()
        title = title.replace(' ','_')
        title = title.lower()
        col_names.append(title)
    '''create table in mysql'''
    sql = 'create table '\
          +table_name+' (' \
          +'id int NOT NULL AUTO_INCREMENT PRIMARY KEY, ' \
          +'at_company varchar(10) DEFAULT \'821\', '

    for i in range(0, ncols):
        sql = sql + col_names[i] + ' varchar(150)'
        if i != ncols-1:
            sql += ','
    sql = sql + ')'
    try:
        cursor.execute(sql)
    except mysql.connector.errors.ProgrammingError as e:
        print e
        # return -1

    '''insert data'''
    #construct sql statement
    sql = 'insert into '+table_name+'('
    for i in range(0, ncols-1):
        sql = sql + col_names[i] + ', '
    sql = sql + col_names[ncols-1]
    sql += ') values ('
    sql = sql + '%s,'*(ncols-1)
    sql += '%s)'
    #get parameters
    parameter_list = []
    for row in xrange(2, nrows):
        for col in range(0, ncols):
            cell_type = sheet.cell_type(row, col)
            cell_value = sheet.cell_value(row, col)
            if cell_type == xlrd.XL_CELL_DATE:
                dt_tuple = xlrd.xldate_as_tuple(cell_value, file.datemode)
                meta_data = str(datetime.datetime(*dt_tuple))
            else:
                meta_data = sheet.cell(row, col).value
            parameter_list.append(meta_data)
        # cursor.execute(sql, parameter_list)
        try:
            cursor.execute(sql, parameter_list)
            parameter_list = []
            ret += 1
        except mysql.connector.errors.ProgrammingError as e:
            print e
        except mysql.connector.errors.DatabaseError as e:
            with open('err_log.txt','a+') as err:
                err.write('[ %s ]文件中第[ %d ]行数据包含特殊字符，插入数据库失败。\n'%(file_path, ret+3))
    return ret



if __name__ == "__main__":
    # if len(sys.argv)<5:
    #     print "Missing Parameters"
    #     sys.exit()
    # elif len(sys.argv)>5:
    #     print "Too Many Parameters"
    #     sys.exit()
    # username = sys.argv[1] posts_from_2013_12_20
    # password = sys.argv[2]
    # database = sys.argv[3]
    # datapath = sys.argv[4]
    username = 'root'
    password = 'root'
    database = 'test3'
    datapath = 'I:\PyProj\importData\data_path'
    importDataHelper(username, password, database, datapath)