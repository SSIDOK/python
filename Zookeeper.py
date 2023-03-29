import os
import win32api
import re

class Keyboard:
    def __init__(self, name, model, type, size, connect):
        self.name = name  # Назва клавіатури
        self.model = model  # Модель клавіатури
        self.type = type  # Тип клавіатури
        self.size = size  # Розмір клавіатури
        self.connect = connect  # Тип підключення клавіатури

    def __str__(self):
        return f"{self.name} {self.model} {self.type}, {self.size}, {self.connect}"  # повертає рядок клавіатури, розділеними комами

    # редагування об'єкта класу Keyboard
    def edit(self, name=None, model=None, type=None, size=None, connect=None):
        if name:
            self.name = name
        if model:
            self.model = model
        if type:
            self.type = type
        if size:
            self.size = size
        if connect:
            self.connect = connect


class KeyboardsFile:
    def __init__(self, keyboards):
        self.file_name = keyboards

    # читає збережені клавіатури з файлу
    def read_keyboards(self):
        try:
            with open(self.file_name, "r", encoding='utf-8') as f:  # відкриття та закриття файлу
                keyboards = [Keyboard(*line.strip().split(", ")) for line in f]
            return keyboards
        except ValueError as e:
            print(f"Сталася помилка: {str(e)}")
        except Exception as e:
            print(f"Сталася невідома помилка{str(e)}")

    #  додає нову клавіатуру у кінець файлу
    def add_keyboard(self, keyboard):
        try:
            with open(self.file_name, "a", encoding='utf-8') as f:
                f.write(f"{keyboard.name}, {keyboard.model}, {keyboard.type}, {keyboard.size}, {keyboard.connect}\n")
        except ValueError as e:
            print(f"Сталася помилка: {str(e)}")
        except Exception as e:
            print(f"Сталася невідома помилка{str(e)}")

    # видаляє останній елемент з файлу
    def delete_last_keyboard(self):
        try:
            keyboards = self.read_keyboards()
            keyboards.pop()
            with open(self.file_name, "w", encoding='utf-8') as f:
                for keyboard in keyboards:
                    f.write(f"{keyboard.name}, {keyboard.model}, {keyboard.type}, {keyboard.size}, {keyboard.connect}\n")
        except ValueError as e:
            print(f"Сталася помилка: {str(e)}")
        except Exception as e:
            print(f"Сталася невідома помилка{str(e)}")

    # змінює дані клавіатури з вказаною назвою
    def update_keyboard(self):
        try:
            keyboards = self.read_keyboards()
            name = input("Введіть назву клавіатури для редагування: ")
            for keyboard in keyboards:
                if keyboard.name == name:
                    new_name = input(f"Введіть нову назву клавіатури (поточне значення: {keyboard.name}): ")
                    new_model = input(f"Введіть нову модель клавіатури (поточне значення: {keyboard.model}):")
                    new_type = input(f"Введіть новий тип клавіатури (поточне значення: {keyboard.type}): ")
                    new_size = input(f"Введіть новий розмір клавіатури (поточне значення: {keyboard.size}): ")
                    new_connect = input(f"Введіть новий тип підключення (поточне значення: {keyboard.connect}): ")
                    keyboard.edit(name=new_name, model=new_model, type=new_type, size=new_size, connect=new_connect)
                    break
            else:
                print(f"Клавіатури з назвою {name} не знайдено")
                return
            with open(self.file_name, "w", encoding='utf-8') as f:
                for keyboard in keyboards:
                    f.write(f"{keyboard.name}, {keyboard.type}, {keyboard.size}, {keyboard.connect}\n")
        except ValueError as e:
            print(f"Сталася помилка {str(e)}")
        except Exception as e:
            print(f"Сталася невідома помилка{str(e)}")

    # сортує клавіатури
    def sort_keyboards(self):
        try:
            keyboards = self.read_keyboards()
            sorted_keyboards = sorted(keyboards, key=lambda k: (k.name, k.type, k.size))
            self.write_keyboards(sorted_keyboards)
            print("Клавіатури відсортовані")
        except ValueError as e:
            print(f"Сталася помилка:{str(e)}")
        except Exception as e:
            print(f"Сталася невідома помилка{str(e)}")

    def write_keyboards(self, keyboards):
        try:
            with open(self.file_name, "w", encoding='utf-8') as f:
                for keyboard in keyboards:
                    f.write(f"{keyboard.name},  {keyboard.model},{keyboard.type}, {keyboard.size}, {keyboard.connect}\n")
        except ValueError as e:
            print(f"Сталася помилка:{str(e)}")
        except Exception as e:
            print(f"Сталася невідома помилка{str(e)}")


class KeyboardMenu:
    def __init__(self, keyboards_file):
        self.keyboards_file = keyboards_file

    # меню
    def display_menu(self):
        while True:
            print(
                "1 - Вивести список клавіатур\n2 - Додати клавіатуру\n3 - Видалити останній елемент\n"
                "4 - Редагувати клавіатуру\n5 - Сортувати клавіатури\n6 - Пустити на друк\n"
                "7 - Знайти клавіатуру за назвою\n8 - Знайти клавіатуру за розміром\n"
                "9 - Пошук клавіатур за першою літерою\n10 - Вийти з програми")
            try:
                choice = int(input("Виберіть дію: "))
                if choice == 1:
                    self.display_keyboards()
                elif choice == 2:
                    self.add_keyboard()
                elif choice == 3:
                    self.delete_last_keyboard()
                elif choice == 4:
                    self.edit_keyboard()
                elif choice == 5:
                    self.sort_keyboards()
                elif choice == 6:
                    self.print_keyboards()
                elif choice == 7:
                    self.search_by_name()
                elif choice == 8:
                    self.search_by_size()
                elif choice == 9:
                    self.search_by_letter()
                elif choice == 10:
                    break
                else:
                    raise ValueError
                input("Натисніть Enter, щоб продовжити...")
                os.system('cls')
            except ValueError:
                print("Введіть число!")
            except Exception as e:
                print(f"Сталася невідома помилка{str(e)}")

        print("Кінець роботи")

    # відображає список клавіатур збережений у файлі
    def display_keyboards(self):
        keyboards = self.keyboards_file.read_keyboards()
        print("Список клавіатур:")
        for i, keyboard in enumerate(keyboards, start=1):
            print(f"{i}. {keyboard}")

    # додаває нові клавіатури до файлу
    def add_keyboard(self):
        try:
            name = input("Введіть назву клавіатури: ")
            type = input("Введіть тип клавіатури: ")
            size = input("Введіть розмір клавіатури: ")
            connect = input("Введіть тип підключення: ")
            keyboard = Keyboard(name, type, size, connect)
            self.keyboards_file.add_keyboard(keyboard)
            print("Записано")
        except ValueError:
            print("Сталася помилка")
        except Exception:
            print("Сталася невідома помилка")

    # видаляє останній елемент з файлу
    def delete_last_keyboard(self):
        self.keyboards_file.delete_last_keyboard()
        print("Видалено останню клавіатуру")

    # сортує клавіатури: за назвою, типом, розміром
    def sort_keyboards(self):
        print("Сортування клавіатур:")
        print("1 - За назвою\n2 - За типом\n3 - За розміром")
        try:
            choice = int(input("Виберіть поле для сортування: "))
            order = input("Введіть порядок сортування (abc/zxy): ")
            if choice == 1:
                key = lambda k: k.name
            elif choice == 2:
                key = lambda k: k.type
            elif choice == 3:
                key = lambda k: -int(re.findall(r'\d+', k.size)[0])
            else:
                raise ValueError

            if order == 'a-z':
                reverse = False
            elif order == 'z-a':
                reverse = True
            else:
                raise ValueError

            keyboards = self.keyboards_file.read_keyboards()
            sorted_keyboards = sorted(keyboards, key=key, reverse=reverse)
            print("Список клавіатур:")
            for i, keyboard in enumerate(sorted_keyboards):
                print(f"{i + 1}. {keyboard}")
        except ValueError:
            print("Неправильне значення введено!")
        except Exception as e:
            print(f"Сталася невідома помилка{str(e)}")

    # редагується наявні клавіатури у файлі
    def edit_keyboard(self):
        keyboards = self.keyboards_file.read_keyboards()
        print("Виберіть клавіатуру для редагування:")
        for i, keyboard in enumerate(keyboards):
            print(f"{i + 1} - {keyboard}")
        try:
            choice = int(input("Ваш вибір: "))
            if choice < 1 or choice > len(keyboards):
                raise ValueError
            keyboard = keyboards[choice - 1]
            name = input(f"Введіть нову назву ({keyboard.name}): ")
            name = name.strip() or keyboard.name
            model = input(f"Введіть нову модель ({keyboard.model}): ")
            model = model.strip() or keyboard.model
            type = input(f"Введіть новий тип ({keyboard.type}): ")
            type = type.strip() or keyboard.type
            size = input(f"Введіть новий розмір ({keyboard.size}): ")
            size = size.strip() or keyboard.size
            connect = input(f"Введіть новий тип підключення ({keyboard.connect}): ")
            connect = connect.strip() or keyboard.connect
            keyboard.edit(name=name, type=type, size=size, connect=connect)
            self.keyboards_file.write_keyboards(keyboards)
            print("Збережено зміни")
        except ValueError:
            print("Неправильний вибір!")
        except Exception as e:
            print(f"Сталася невідома помилка{str(e)}")

    # виводить файл з клавіатурами на друк
    def print_keyboards(self):
        try:
            keyboards = self.keyboards_file.read_keyboards()
            print("Список клавіатур:")
            for keyboard in keyboards:
                print(keyboard)

            file_name = input("Введіть назву файлу для друку: ")
            if not file_name.endswith(".txt"):
                file_name += ".txt"

            with open(file_name, "w") as f:
                for keyboard in keyboards:
                    f.write(f"{keyboard.name}, {keyboard.type}, {keyboard.size}, {keyboard.connect}\n")

            win32api.ShellExecute(0, "print", file_name, None, ".", 0)
        except ValueError as e:
            print(f"Сталася помилка{str(e)}")
        except Exception as e:
            print(f"Сталася невідома помилка{str(e)}")

    # пошук за назвою клавіатури
    def search_by_name(self):
        while True:
            try:
                keyboards = self.keyboards_file.read_keyboards()
                name = input("Введіть назву клавіатури: ")
                found_keyboards = [keyboard for keyboard in keyboards if keyboard.name == name]
                if not found_keyboards:
                    print("Клавіатур з такою назвою не знайдено")
                else:
                    for keyboard in found_keyboards:
                        print(keyboard)
                break  # вихід з циклу while, якщо введено правильне значення
            except ValueError as e:
                print(f"Сталася помилка: {str(e)}")
            except Exception as e:
                print(f"Сталася невідома помилка: {str(e)}")

    # пошук за розміром клавіатури
    def search_by_size(self):
        while True:
            try:
                keyboards = self.keyboards_file.read_keyboards()
                size = input("Введіть розмір клавіатури: ")
                found_keyboards = [keyboard for keyboard in keyboards if keyboard.size == size]
                if not found_keyboards:
                    print("Клавіатур з таким розміром не знайдено")
                else:
                    for keyboard in found_keyboards:
                        print(keyboard)
                    break  # вихід з циклу while, якщо вдалося знайти клавіатуру
            except ValueError as e:
                print(f"Сталася помилка введення: {str(e)}")
            except Exception as e:
                print(f"Сталася невідома помилка: {str(e)}")

    # пошук клавіатур за першою літерою
    def search_by_letter(self):
        try:
            keyboards = self.keyboards_file.read_keyboards()
            while True:
                letter = input("Введіть літеру за якою буде виконуватися пошук: ")
                if not letter.isalpha() or len(letter) != 1:
                    print("Введіть одну літеру")
                else:
                    found_keyboards = [keyboard for keyboard in keyboards if keyboard.name.startswith(letter)]
                    if not found_keyboards:
                        print("Клавіатур не знайдено")
                    else:
                        for keyboard in found_keyboards:
                            print(keyboard)
                        break  # вийти з циклу while, якщо знайдено клавіатури
        except Exception as e:
            print(f"Сталася помилка: {str(e)}")


if __name__ == '__main__':
    file_name = "keyboards.txt"
    keyboards_file = KeyboardsFile(file_name)
    menu = KeyboardMenu(keyboards_file)
    menu.display_menu()
