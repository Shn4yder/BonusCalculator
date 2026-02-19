import aspose.tasks as tsk
import sys
import os
import shutil

# Путь к исходному файлу
source_mpp = r"C:\Users\a.komarkova\Documents\projects\BonusCalculator\data\input\input_project.mpp"
temp_mpp = r"C:\Users\a.komarkova\Documents\projects\BonusCalculator\data\input\temp_test_project.mpp"

def main():
    if not os.path.exists(source_mpp):
        print(f"Файл не найден: {source_mpp}")
        return

    # Копируем файл во временный, чтобы избежать блокировок
    try:
        shutil.copyfile(source_mpp, temp_mpp)
    except Exception as e:
        print(f"Ошибка копирования файла: {e}")
        return

    try:
        print(f"Открываем проект: {temp_mpp}")
        project = tsk.Project(temp_mpp)
        
        # Вывод имени корневой задачи
        if project.root_task:
            print(f"ROOT TASK NAME: '{project.root_task.name}'")
        else:
            print("ROOT TASK is None")
            
        print("-" * 30)
        print("Список всех задач:")
        
        # Рекурсивная функция для вывода дерева задач
        def print_task_tree(task, level=0):
            indent = "  " * level
            name = task.name if task.name else "[No Name]"
            print(f"{indent}- {name} (ID: {task.id})")
            
            for child in task.children:
                print_task_tree(child, level + 1)
                
        # Запускаем обход с корневой задачи
        if project.root_task:
             # Если хотим видеть и сам корень в списке:
             print_task_tree(project.root_task)
        
        # Альтернативный способ: плоский список всех задач
        # print("-" * 30)
        # print("Плоский список (через select_all_child_tasks):")
        # for t in project.root_task.select_all_child_tasks():
        #     print(f"ID: {t.id}, Name: {t.name}")

    except Exception as e:
        print(f"Ошибка при чтении проекта: {e}")
    finally:
        # Удаляем временный файл
        if os.path.exists(temp_mpp):
            try:
                os.remove(temp_mpp)
            except:
                pass

if __name__ == "__main__":
    main()
