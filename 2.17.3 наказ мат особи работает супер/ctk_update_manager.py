# ctk_update_manager.py
"""
Простий менеджер оновлень для CustomTkinter
Спрощена версія без складних потоків
"""

import customtkinter as ctk
import tkinter.messagebox as messagebox
import threading
import time
from updater import AutoUpdater, check_updates


class SimpleUpdateManager:
    """Простий менеджер оновлень для CustomTkinter програм"""

    def __init__(self, parent_window: ctk.CTk, repo_owner: str, repo_name: str, current_version: str):
        self.parent = parent_window
        self.repo_owner = repo_owner
        self.repo_name = repo_name
        self.current_version = current_version.lstrip('v')

        # Автоперевірка через 5 секунд після запуску
        self.parent.after(5000, self.auto_check_updates)

    def auto_check_updates(self):
        """Автоматична перевірка оновлень у фоні"""

        def check_background():
            try:
                update_info = check_updates(self.repo_owner, self.repo_name, self.current_version)
                if update_info and update_info.get('newer', False):
                    # Показуємо діалог в основному потоці
                    self.parent.after(0, lambda: self.show_update_notification(update_info))
            except Exception as e:
                print(f"[AUTO UPDATE] Помилка автоперевірки: {e}")

        threading.Thread(target=check_background, daemon=True).start()

    def check_updates_manual(self):
        """Ручна перевірка оновлень з простим діалогом"""
        # Простий діалог перевірки
        check_dialog = ctk.CTkToplevel(self.parent)
        check_dialog.title("Перевірка оновлень")
        check_dialog.geometry("300x150")
        check_dialog.resizable(False, False)
        check_dialog.transient(self.parent)
        check_dialog.grab_set()

        # Центрування
        self._center_window(check_dialog, 300, 150)

        # Контент
        status_label = ctk.CTkLabel(check_dialog, text="🔍 Перевірка оновлень...",
                                    font=ctk.CTkFont(size=16))
        status_label.pack(pady=30)

        progress = ctk.CTkProgressBar(check_dialog, mode="indeterminate")
        progress.pack(pady=10, padx=30, fill="x")
        progress.start()

        cancel_btn = ctk.CTkButton(check_dialog, text="Скасувати", width=100,
                                   command=check_dialog.destroy)
        cancel_btn.pack(pady=10)

        def check_thread():
            try:
                time.sleep(1)  # Мінімальна затримка для UX
                update_info = check_updates(self.repo_owner, self.repo_name, self.current_version)

                # Закриваємо діалог перевірки
                self.parent.after(0, check_dialog.destroy)

                # Показуємо результат
                if update_info and update_info.get('newer', False):
                    self.parent.after(0, lambda: self.show_update_notification(update_info))
                else:
                    self.parent.after(0, self.show_no_updates)

            except Exception as e:
                self.parent.after(0, check_dialog.destroy)
                self.parent.after(0, lambda: self.show_error(f"Помилка перевірки: {str(e)}"))

        threading.Thread(target=check_thread, daemon=True).start()

    def show_update_notification(self, update_info: dict):
        """Показує сповіщення про доступне оновлення"""
        # Створюємо діалог оновлення
        update_dialog = ctk.CTkToplevel(self.parent)
        update_dialog.title("🚀 Доступне оновлення")
        update_dialog.geometry("450x350")
        update_dialog.resizable(False, False)
        update_dialog.transient(self.parent)
        update_dialog.grab_set()

        self._center_window(update_dialog, 450, 350)

        # Заголовок
        title = ctk.CTkLabel(update_dialog, text="🚀 Нова версія доступна!",
                             font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=20)

        # Інформація про версії
        info_frame = ctk.CTkFrame(update_dialog)
        info_frame.pack(pady=15, padx=20, fill="x")

        current_lbl = ctk.CTkLabel(info_frame, text=f"📦 Поточна: v{self.current_version}",
                                   font=ctk.CTkFont(size=14))
        current_lbl.pack(pady=5)

        new_lbl = ctk.CTkLabel(info_frame, text=f"🔄 Нова: v{update_info['version']}",
                               font=ctk.CTkFont(size=14, weight="bold"),
                               text_color="green")
        new_lbl.pack(pady=5)

        # Примітки релізу
        notes_lbl = ctk.CTkLabel(update_dialog, text="📝 Що нового:",
                                 font=ctk.CTkFont(size=14, weight="bold"))
        notes_lbl.pack(pady=(15, 5))

        notes_text = ctk.CTkTextbox(update_dialog, height=100, width=400)
        notes_text.pack(pady=5, padx=20)

        release_notes = update_info.get('release_notes', 'Немає опису')
        if len(release_notes) > 400:
            release_notes = release_notes[:400] + "..."
        notes_text.insert("1.0", release_notes)
        notes_text.configure(state="disabled")

        # Кнопки
        btn_frame = ctk.CTkFrame(update_dialog)
        btn_frame.pack(pady=20, padx=20, fill="x")

        update_btn = ctk.CTkButton(btn_frame, text="🔄 Оновити зараз",
                                   command=lambda: self.start_update(update_dialog),
                                   width=150, font=ctk.CTkFont(size=14, weight="bold"))
        update_btn.pack(side="left", padx=10)

        cancel_btn = ctk.CTkButton(btn_frame, text="❌ Пізніше",
                                   command=update_dialog.destroy, width=150,
                                   fg_color="gray", hover_color="darkgray")
        cancel_btn.pack(side="right", padx=10)

    def start_update(self, notification_dialog):
        """Запускає процес оновлення"""
        notification_dialog.destroy()

        # Створюємо вікно прогресу
        progress_dialog = ctk.CTkToplevel(self.parent)
        progress_dialog.title("Оновлення SportForAll")
        progress_dialog.geometry("400x250")
        progress_dialog.resizable(False, False)
        progress_dialog.transient(self.parent)
        progress_dialog.grab_set()

        self._center_window(progress_dialog, 400, 250)

        # Заголовок
        title = ctk.CTkLabel(progress_dialog, text="🚀 Оновлення SportForAll",
                             font=ctk.CTkFont(size=18, weight="bold"))
        title.pack(pady=20)

        # Статус
        status_label = ctk.CTkLabel(progress_dialog, text="Підготовка...",
                                    font=ctk.CTkFont(size=14))
        status_label.pack(pady=10)

        # Прогрес бар
        progress_bar = ctk.CTkProgressBar(progress_dialog, width=350)
        progress_bar.pack(pady=15, padx=25)
        progress_bar.set(0)

        # Відсотки
        percent_label = ctk.CTkLabel(progress_dialog, text="0%")
        percent_label.pack(pady=5)

        # Лог
        log_frame = ctk.CTkFrame(progress_dialog)
        log_frame.pack(pady=10, padx=20, fill="both", expand=True)

        log_text = ctk.CTkTextbox(log_frame, height=60, font=ctk.CTkFont(size=10))
        log_text.pack(pady=10, padx=10, fill="both", expand=True)

        def add_log(message):
            def update_gui():
                if progress_dialog.winfo_exists():
                    log_text.insert("end", f"{message}\n")
                    log_text.see("end")
                    progress_dialog.update_idletasks()

            self.parent.after(0, update_gui)

        def update_progress(percent, status):
            def update_gui():
                if progress_dialog.winfo_exists():
                    status_label.configure(text=status)
                    progress_bar.set(percent / 100.0)
                    percent_label.configure(text=f"{percent:.1f}%")
                    progress_dialog.update_idletasks()

            self.parent.after(0, update_gui)

        def update_thread():
            try:
                add_log("Початок оновлення...")
                updater = AutoUpdater(self.repo_owner, self.repo_name, self.current_version)

                result = updater.perform_full_update(update_progress)

                # Результат оновлення
                self.parent.after(0, lambda: self.handle_update_result(result, progress_dialog))

            except Exception as e:
                error_msg = f"Критична помилка: {str(e)}"
                add_log(error_msg)
                self.parent.after(0, lambda: self.handle_update_error(error_msg, progress_dialog))

        add_log("Запуск процесу оновлення...")
        threading.Thread(target=update_thread, daemon=True).start()

    def handle_update_result(self, result: dict, progress_dialog):
        """Обробляє результат оновлення"""
        progress_dialog.destroy()

        if result['success']:
            messagebox.showinfo(
                "✅ Оновлення завершено",
                f"Оновлення до версії {result['version']} встановлено!\n\n"
                "🔄 Програма зараз перезапуститься для застосування змін.",
                parent=self.parent
            )
            # Закриваємо програму для перезапуску
            self.parent.quit()
        else:
            if result.get('no_update'):
                messagebox.showinfo("ℹ️ Інформація", result['message'], parent=self.parent)
            else:
                error_msg = result.get('error', 'Невідома помилка')
                messagebox.showerror(
                    "❌ Помилка оновлення",
                    f"Не вдалося оновити програму:\n\n{error_msg}\n\n"
                    "💡 Спробуйте:\n"
                    "• Запустити програму як адміністратор\n"
                    "• Перевірити антивірус\n"
                    "• Завантажити вручну з GitHub",
                    parent=self.parent
                )

    def handle_update_error(self, error_msg: str, progress_dialog):
        """Обробляє помилку оновлення"""
        progress_dialog.destroy()
        messagebox.showerror(
            "❌ Критична помилка",
            f"Сталася помилка при оновленні:\n\n{error_msg}",
            parent=self.parent
        )

    def show_no_updates(self):
        """Показує повідомлення що оновлень немає"""
        messagebox.showinfo(
            "✅ Актуальна версія",
            f"У вас остання версія програми!\n\n📦 Версія: v{self.current_version}",
            parent=self.parent
        )

    def show_error(self, error_msg: str):
        """Показує помилку"""
        messagebox.showerror(
            "❌ Помилка",
            f"Не вдалося перевірити оновлення:\n\n{error_msg}\n\n"
            "Перевірте з'єднання з Інтернетом.",
            parent=self.parent
        )

    def _center_window(self, window, width, height):
        """Центрує вікно на екрані"""
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()

        x = (screen_width - width) // 2
        y = (screen_height - height) // 2

        window.geometry(f"{width}x{height}+{x}+{y}")


# Проста функція для налаштування оновлень
def setup_auto_updates(main_window: ctk.CTk, current_version: str) -> SimpleUpdateManager:
    """
    Простий спосіб додати автооновлення до SportForAll

    Використання:
    from ctk_update_manager import setup_auto_updates

    update_manager = setup_auto_updates(main_window, "2.16.4")

    # Для ручної перевірки додайте кнопку:
    # check_btn = ctk.CTkButton(main_window, text="Перевірити оновлення",
    #                          command=update_manager.check_updates_manual)
    """
    return SimpleUpdateManager(
        main_window,
        "KostyaBulavko",  # твоє GitHub ім'я
        "sportfprall",  # назва репозиторію
        current_version
    )


# Приклад використання
if __name__ == "__main__":
    # Тест GUI
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Тест оновлень")
    root.geometry("300x200")

    update_manager = setup_auto_updates(root, "2.16.3")

    # btn = ctk.CTkButton(root, text="Перевірити оновлення",
    #                     command=update_manager.check_updates_manual)
    # btn.pack(pady=50)

    root.mainloop()