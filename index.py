import pandas as pd

def create_teacher_ranking_excel(input_excel_path, output_excel_path='рейтинг_вчителів.xlsx'):
    """
    Розраховує рейтинг вчителів за середнім балом їхніх оцінок
    з різних листів Excel-файлу та зберігає його у новий Excel-файл.

    Args:
        input_excel_path (str): Шлях до вхідного Excel-файлу з оцінками.
        output_excel_path (str): Шлях для збереження вихідного Excel-файлу з рейтингом.
                                 За замовчуванням: 'рейтинг_вчителів.xlsx'.

    Returns:
        bool: True, якщо рейтинг успішно створено та збережено, False у разі помилки.
    """
    try:
        # 1. Читаємо всі листи Excel-файлу
        # sheet_name=None читає всі листи у словник, де ключ - назва листа, значення - DataFrame.
        all_sheets_data = pd.read_excel(input_excel_path, sheet_name=None, header=None)
        # header=None важливо, бо ми самі вручну візьмемо імена вчителів з першого рядка
    except FileNotFoundError:
        print(f"Помилка: Вхідний файл '{input_excel_path}' не знайдено.")
        return False
    except Exception as e:
        print(f"Помилка при читанні вхідного файлу Excel: {e}")
        return False

    all_teachers_scores = {} # Словник для зберігання всіх оцінок кожного вчителя

    # 2. Обробляємо кожен лист
    for sheet_name, df in all_sheets_data.items():
        if df.empty:
            print(f"Попередження: Лист '{sheet_name}' порожній. Пропущено.")
            continue

        # Перший рядок DataFrame - це імена вчителів.
        # Вибираємо перший рядок і конвертуємо його в список.
        # .iloc[0] повертає Series, тому .tolist() працює коректно.
        teacher_names_on_sheet = df.iloc[0].tolist()

        # Решта рядків (починаючи з індексу 1) - це оцінки.
        # Вибираємо їх, ігноруючи перший рядок.
        scores_only_df = df.iloc[1:].copy()

        # Тепер нам потрібно призначити імена вчителів як заголовки колонок для оцінок.
        # Оскільки teacher_names_on_sheet може містити NaN або порожні рядки,
        # перетворимо їх на валідні імена (наприклад, 'Вчитель_без_імені_N')
        # або просто усунемо NaN, якщо вони є.
        # Найбезпечніше - використовувати оригінальні імена, але переконатись, що вони не NaN
        cleaned_teacher_names = [
            name if pd.notna(name) and str(name).strip() != '' else f'Unnamed_Teacher_{i}'
            for i, name in enumerate(teacher_names_on_sheet)
        ]
        scores_only_df.columns = cleaned_teacher_names

        # Проходимося по кожному вчителю на поточному листі
        for teacher_name in cleaned_teacher_names:
            # Отримуємо колонку оцінок для цього вчителя
            # .dropna() видаляє будь-які NaN (порожні клітинки) з цієї колонки
            # .tolist() конвертує Series у звичайний Python список
            current_teacher_scores = scores_only_df[teacher_name].dropna().tolist()

            # Додаємо ці оцінки до загального списку оцінок цього вчителя
            if teacher_name not in all_teachers_scores:
                all_teachers_scores[teacher_name] = []
            all_teachers_scores[teacher_name].extend(current_teacher_scores)

    if not all_teachers_scores:
        print("Не знайдено жодних вчителів чи оцінок для розрахунку рейтингу.")
        return False

    # 3. Розраховуємо середній бал для кожного вчителя
    teacher_avg_scores = []
    for teacher, scores in all_teachers_scores.items():
        if scores:
            average = sum(scores) / len(scores)
            teacher_avg_scores.append({'Вчитель': teacher, 'Середній Бал': average})
        else:
            teacher_avg_scores.append({'Вчитель': teacher, 'Середній Бал': 0}) # Вчитель без оцінок

    # 4. Створюємо DataFrame для рейтингу
    rating_df = pd.DataFrame(teacher_avg_scores)

    # 5. Сортуємо за середнім балом у спадному порядку
    rating_df = rating_df.sort_values(by='Середній Бал', ascending=False).reset_index(drop=True)

    # 6. Додаємо колонку з рейтингом (позицією)
    rating_df['Рейтинг'] = rating_df.index + 1

    # 7. Зберігаємо рейтинг у новий Excel-файл
    try:
        rating_df.to_excel(output_excel_path, index=False)
        print(f"\nРейтинг успішно створено та збережено у файл: '{output_excel_path}'")
        return True
    except Exception as e:
        print(f"Помилка при збереженні вихідного файлу Excel: {e}")
        return False

# --- Приклад використання функції ---
if __name__ == "__main__":
    input_file = 'оцінки_вчителів.xlsx' # Переконайтесь, що цей файл існує у тій же папці або вкажіть повний шлях.
    output_file = 'рейтинг_вчителів_за_середнім.xlsx'

    success = create_teacher_ranking_excel(input_file, output_file)

    if success:
        # Для перевірки можна прочитати створений файл
        print("\nВміст створеного файлу:")
        try:
            generated_rating = pd.read_excel(output_file)
            print(generated_rating)
        except Exception as e:
            print(f"Не вдалося прочитати щойно створений файл: {e}")
    else:
        print("Не вдалося створити рейтинг.")