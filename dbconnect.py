import pymysql
import xlrd

f: xlrd.book.Book = xlrd.open_workbook_xls('IMDB_Reviews.xls')

table1: xlrd.sheet.Sheet = f.sheets()[0]

table2: xlrd.sheet.Sheet = f.sheets()[1]

IS_INITIAL = True

db = pymysql.connect(host='192.168.1.166',
                     port=3306,
                     user='lisn',
                     passwd='123456',
                     database='lisn')

cursor = db.cursor()

def execute_sql(sql):
    try:
        cursor.execute(sql)
        db.commit()
    except:
        db.rollback()
        print('error!')

if IS_INITIAL:
    sql = " CREATE TABLE MOVIE_DETAILS \
        (MOVIE_ID INT, \
        TITLE TEXT, \
        GENRE CHAR(20), \
        DIRECTORS CHAR(50), \
        WRITERS TEXT, \
        STARS TEXT, \
        STORYLINE TEXT) "

    cursor.execute(sql)

    sql = " CREATE TABLE MOVIE_REVIEWS \
        (REVIEW_ID INT, \
        MOVIE_ID INT, \
        TITLE TEXT, \
        AUTHOR TEXT, \
        DATE CHAR(50), \
        UP_VOTE INT, \
        TOTAL_VOTE INT, \
        RATING CHAR(10), \
        REVIEW TEXT) "

    cursor.execute(sql)

for i in range(1, table1.nrows):
    
    row = table1.row_values(i)
    row = [txt.replace('\'', '`') if isinstance(txt, str) else txt for txt in row]
    
    movie_id, title, genre, director, writers, stars, storyline = row

    sql = " INSERT INTO MOVIE_DETAILS \
            (MOVIE_ID, TITLE, GENRE, DIRECTORS, WRITERS, STARS, STORYLINE) \
            VALUES( \
            %s, '%s', '%s', '%s', '%s', '%s', '%s')" % (movie_id, title, genre, director, writers, stars, storyline)
    execute_sql(sql)

for i in range(1, table2.nrows):
    
    row = table2.row_values(i)
    row = [txt.replace('\'', '`') if isinstance(txt, str) else txt for txt in row]
    
    review_id, movie_id, title, author, date, upvote, totalvote, rating, review = row

    sql = " INSERT INTO MOVIE_REVIEWS( \
            REVIEW_ID, MOVIE_ID, TITLE, AUTHOR, DATE, UP_VOTE, TOTAL_VOTE, RATING, REVIEW)\
            VALUES( \
            %s, %s, '%s', '%s', '%s', %s, %s, '%s', '%s')" % (review_id, movie_id, title, author, date, int(upvote), int(totalvote), rating, review)
    execute_sql(sql)


db.close()