import os
import win32com.client as win32

# Запрос корневой папки и пароля
root_folder = input('Введите путь к корневой папке: ')
password = 'test'

# Получение списка файлов в корневой папке и ее подпапках
for root, dirs, files in os.walk(root_folder):
    for file in files:
        # Определение расширения файла
        file_extension = os.path.splitext(file)[1]
        # Если файл имеет расширение .docx или .doc
        if file_extension in ('.docx', '.doc'):
            # Формирование полного пути к файлу
            file_path = os.path.join(root, file)
            try:
                # Открытие документа в режиме редактирования
                word = win32.gencache.EnsureDispatch('Word.Application')
                doc = word.Documents.Open(file_path, Password=password, WritePassword=password)
                # Внесение изменений в документ
                doc.Content.Text = 'Документ защищен паролем'
                # Сохранение изменений и закрытие документа
                doc.Save()
                doc.Close()
                # Вывод сообщения об успешном изменении файла
                print(f'Файл {file_path} успешно изменен')
            except Exception as e:
                # Вывод сообщения об ошибке при изменении файла
                print(f'Ошибка при редактировании файла {file_path}: {e}')
            finally:
                # Выход из режима редактирования
                word.Quit()
