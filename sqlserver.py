import pyodbc
import sys

class SQLSVR:
    # def __init__(self, host, user, pwd, db):
    def __init__(self, *conn_string):
        #connect string for connecting to sqlserver 2005
        # self.conn_info = 'DRIVER={SQL Server};DATABASE=%s;SERVER=%s;UID=%s;PWD=%s'%(db, host, user, pwd)
        
        #connect string for connecting to sqlserver 2012
        # self.conn_info = 'DRIVER={SQL Server Native Client 11.0};DATABASE=%s;SERVER=%s;UID=%s;PWD=%s'%(db, host, user, pwd)
        self.conn_info = 'DRIVER={SQL Server Native Client 11.0};DATABASE=%s;SERVER=%s;UID=%s;PWD=%s'%(conn_string[3], conn_string[0], conn_string[1], conn_string[2])

    def __GetConnect(self):
        try:
            self.conn = pyodbc.connect(self.conn_info)
        except pyodbc.DatabaseError:
            print('SQLSERVER CONNECT TO'+self.host+' ERROR')
            sys.exit(0)
        cur = self.conn.cursor()
        if not cur:
            print('SQLSERVER CONNECT TO '+self.host+' ERROR')
        else:
            return(cur)

    def execSql(self, sql):
        cur = self.__GetConnect()
        cur.execute(sql)
        self.conn.commit()
        self.conn.close()

    def querySql(self, sql):
        cur = self.__GetConnect()
        cur.execute(sql)

        index= cur.description
        lists = cur.fetchall()
        resList = []
        # print(len(index))
        row = {}
        for tempreslist in lists:
            for i in range(0, len(index)):
                row[index[i][0]] = tempreslist[i]
                # print(tempreslist[i])
                resList.append(row)
        self.conn.close()
        return(resList)

    def close(self):
        self.conn.close()

if __name__ == '__main__':
    #'luckymyj-Pc', 'sa', 'sa', 'dtpbiz' is for sqlserver2005
    # tmpsql = SQLSVR('luckymyj-Pc', 'sa', 'sa', 'dtpbiz')

    #'LUCKYMYJ-PC\SQLSERVER2012', user = 'sa', pwd = 'myj810714' is for sqlserver2012
    conn_string = ['LUCKYMYJ-PC\SQLSERVER2012', 'sa', 'myj810714', 'courseinfo']
    # sql = SQLSVR(host = 'LUCKYMYJ-PC\SQLSERVER2012', user = 'sa', pwd = 'myj810714', db = 'courseinfo')
    sql = SQLSVR(*conn_string)
    insert_course_string = "insert into courseinfos (id, coursename, people, courseinstitute, courseprice) values"
    delete_course_string = "delete from courseinfos"
    a = 'abc'
    b = 'cbd'
    c = 'bcd'
    d = 'def'
    value = "(%d, \'%s\', \'%s\', \'%s\', \'%s\')"%(1, a, b, c, d)
    print(insert_course_string+value)
    sql.execSql(delete_course_string)
    sql.execSql(insert_course_string+value)
    
    print(sql.querySql('select * from courseinfos'))