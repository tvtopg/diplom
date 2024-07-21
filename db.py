import pymysql

from config import host, user, password, db_name


def databaseQuery ():
    try:
        connection = pymysql.connect( #подключение к базе
            host=host,
            port=3306,
            user=user,
            password=password,
            database=db_name,
            cursorclass=pymysql.cursors.DictCursor
        )
        print("successfully connected") #успешно подключен 
        print("#" * 20)

        try: # Запросы к БД
            # Объект который содержит в себе различные методы для проведения sql команд
            # cursor = connection.cursor()
            # или

            # create table - создание таблицы
            # with connection.cursor() as cursor: # контекстный менеджер (объект который содеожит в себе различные методы для проедения команд)
                # create_table_query = "CREATE TABLE `users`(id int AUTO_INCREMENT," \
                #                      "name varchar(32)," \
                #                      "password varchar(32),"\
                #                      "email varchar(32), PRIMARY KEY (id));"
            #     cursor.execute(create_table_query) # Для создания таблици вызываетм метод execute у курсора
            #     print("Table created successfully")

            # insert data  - добавление данных в таблицу
            # with connection.cursor() as cursor: 
            #     insert_query = "INSERT INTO `users` (name, password, email) VALUES ('Anna', 'qwerty', qwew@gmail.com);"
            #     cursor.execute(create_table_query)
            #     connection.commit() # сохранение 

            # udate data 
            # with connection.cursor() as cursor: 
            #     udate_query = "UPDATE `disciplines` SET disciplineName = 'Философия' WHERE uCode = '09.03.02-Б1.В.9';"
            #     cursor.execute(udate_query)
            #     # connection.commit()

            # delete data 
            # with connection.cursor() as cursor: 
            #     delete_query = "DELETE FROM `disciplines` WHERE uCode = '09.03.02-Б1.В.9';"
            #     cursor.execute(delete_query)
                # connection.commit()

            # drop table # НЕ ТРЕБУЕТ КОММИТА ДЛЯ СОХРАНЕНИЯ
            # with connection.cursor() as cursor: 
            #     drop_table_query = "DROP TABLE `disciplines`;"
            #     cursor.execute(drop_table_query)
            

            # select all data from table
            # with connection.cursor() as cursor: 
            #     select_all_rows = "SELECT * FROM `disciplines`"
            #     cursor.execute(select_all_rows)
            #     rows = cursor.fetchall() # извлекает все строки
            #     for row in rows: # пробегаеся по строкам и выводи
            #         print(row)

            # select inner join
            with connection.cursor() as cursor: 
                select_inner_join = \
                    "SELECT approver, approverName, disciplineName, typeName, directionsSpecialtiesName, typeProgram , programName, levelName, formName, overallComplexityDiscipline \
                    FROM disciplines \
                    INNER JOIN educational_programs ON disciplines.educationalProgramsID = educational_programs.id \
                    INNER JOIN form_education ON educational_programs.formEducationID = form_education.id \
                    INNER JOIN directions_specialties ON educational_programs.directionsSpecialtiesID = directions_specialties.id \
                    INNER JOIN level_education_qualification ON directions_specialties.levelEducationQualificationID = level_education_qualification.id \
                    INNER JOIN faculties_institutes ON directions_specialties.facultiesInstitutesID = faculties_institutes.id;"
                cursor.execute(select_inner_join)
                rows = cursor.fetchall() # извлекает все строки
                # for row in rows: # пробегаеся по строкам и выводи
                #     print(row['approver'])
                #     print(row)



        finally:
            connection.close()  # закрытие соединения
    except Exception as ex:
        print("Connection refused")
        print(ex)
    
    return rows


print (databaseQuery())
