There are two dependency modules need to be installed.
    1. xlrd # to read excel files
    2. mysql-connector-python # to work with Mysql

Directory Structure:
    data_path: test files
    demos: three demos for different steps in the main program
    ImportDataProgram.py: the main program

Procedure:
    (1) Get all the paths and names of the files need to be stored
    (2) Connect MySQL
    (3) Create tables for each file
    (4) Insert data into each table

Usage:
    0. create a new database in mysql
        For example, after logging in mysql in terminal, you can use the the following command
        "create database test_database" to create a database named 'test_database',
        you can replace "test_database" with any other names you like.

    1. set username, password, database(created in step 0) and datapath in the tail of ImportDataProgram.py
    2. run ImportDataProgram.py with the following command
        python ImportDataProgram.py [username] [password] [database] [datapath]
    # username: your username in your mysql
    # password: the corresponding password
    # database: the database you specific
    # datapath: the directory of excel files
    e.g.
        python ImportDataProgram.py root root test_database data_path

PS:
    (1) The Length of Data In Table
    I am not sure the maximum length of data, so I set the
    length of data in mysql tables is 150 characters (you can find
    the code in function storeData(file_path, table_name, cursor), the code is
    " sql = sql + col_names[i] + ' varchar(150)' "), you can adjust it according
     to your requirements.
    (2)Table Name:
     You can set the rules of table name, the code is following the comment code:
    '''convert file name to mysql table name''' in function getFilesList(dir).