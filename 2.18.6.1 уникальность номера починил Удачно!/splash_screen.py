import tkinter as tk
from tkinter import ttk
import customtkinter as ctk
import time
import threading
import sys
from globals import version,name_prog

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")


class SplashScreen:
    def __init__(self, duration=5, app_password="qwer"):
        self.duration = duration
        self.app_password = app_password
        self.success_callback = None
        self.root = tk.Tk()
        self.root.attributes("-alpha", 0.0)  # Повна прозорість при старті

        self.root.title("Завантаження програми...")

        # Налаштування вікна
        window_width = 450
        window_height = 400

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.resizable(False, False)
        self.root.overrideredirect(True)
        self.root.configure(bg = '#4CAF50')

        border_frame = tk.Frame(self.root, bg='#4CAF50', bd=2)
        border_frame.pack(fill='both', expand=True, padx=2, pady=2)

        inner_frame = tk.Frame(border_frame, bg='#1a1a1a')
        inner_frame.pack(fill='both', expand=True, padx=2, pady=2)

        self.create_widgets(inner_frame)

        self.root.attributes('-topmost', True)
        self.root.protocol("WM_DELETE_WINDOW", self.on_window_close)

        self.fade_in()

    def create_widgets(self, parent):
        main_frame = tk.Frame(parent, bg='#1a1a1a')
        main_frame.pack(expand=True, fill='both', padx=30, pady=30)

        tk.Label(main_frame, text="🏃" + name_prog + version,
                 font=('Arial', 18, 'bold'), fg='#4CAF50', bg='#1a1a1a').pack(pady=(0, 8))

        self.subtitle_label = tk.Label(main_frame, text="Завантаження модулів та ініціалізація...",
                                       font=('Arial', 11), fg='#cccccc', bg='#1a1a1a')
        self.subtitle_label.pack(pady=(0, 20))

        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor='#2d2d2d', background='#4CAF50', relief='flat')

        self.progress = ttk.Progressbar(main_frame, style="Custom.Horizontal.TProgressbar",
                                        length=380, mode='determinate')
        self.progress.pack(pady=(0, 10))

        stats_frame = tk.Frame(main_frame, bg='#1a1a1a')
        stats_frame.pack(fill='x', pady=(0, 8))

        self.percent_label = tk.Label(stats_frame, text="0%", font=('Arial', 14, 'bold'),
                                      fg='#4CAF50', bg='#1a1a1a')
        self.percent_label.pack(side='left')

        self.time_label = tk.Label(stats_frame, text=f"Залишилось: {self.duration} сек",
                                   font=('Arial', 10), fg='#888888', bg='#1a1a1a')
        self.time_label.pack(side='right')

        self.status_label = tk.Label(main_frame, text="Підготовка до запуску...",
                                     font=('Arial', 9), fg='#666666', bg='#1a1a1a')
        self.status_label.pack(pady=(5, 15))

        # Блок пароля
        self.password_frame = tk.Frame(main_frame, bg='#1a1a1a')
        self.password_frame.pack_forget()

        tk.Frame(self.password_frame, height=1, bg='#333333').pack(fill='x', pady=(0, 15))



        password_entry_frame = tk.Frame(self.password_frame, bg='#1a1a1a')
        password_entry_frame.pack(pady=(0, 10))

        # 👇 Сучасне поле з customtkinter
        self.password_entry = ctk.CTkEntry(
            master=password_entry_frame,
            show="*",
            width=220,
            height=40,
            font=('Arial', 16),
            justify='center',
            text_color="white",
            placeholder_text="Введіть пароль",
            placeholder_text_color="#888888",
            fg_color="#1e1e1e",  # темно-сірий фон
            border_color="#2d2d2d",  # тонка рамка майже невидима
            border_width=1,
            corner_radius=4  # майже прямокутне
        )

        self.password_entry.pack(padx=5, pady=5)

        # Кнопки
        buttons_frame = tk.Frame(self.password_frame, bg='#1a1a1a')
        buttons_frame.pack(pady=(5, 0))

        self.login_button = ctk.CTkButton(
            master=buttons_frame,
            text="Увійти",
            command=self.check_password,
            fg_color="#4CAF50",
            hover_color="#45a049",
            text_color="white",
            font=('Arial', 12, 'bold'),
            width=100
        )
        self.login_button.pack(side='left', padx=(0, 10))

        self.exit_button = ctk.CTkButton(
            master=buttons_frame,
            text="Вихід",
            command=self.on_window_close,
            fg_color="#f44336",
            hover_color="#e53935",
            text_color="white",
            font=('Arial', 12),
            width=100
        )
        self.exit_button.pack(side='left')

        self.error_label = tk.Label(self.password_frame, text="",
                                    font=('Arial', 9), fg='#f44336', bg='#1a1a1a')
        self.error_label.pack(pady=(5, 0))

        self.main_frame = main_frame

    def check_password(self):
        entered_password = self.password_entry.get()
        if entered_password == self.app_password:
            self.error_label.config(text="✓ Пароль правильний!", fg='#4CAF50')
            self.root.after(500, self.close_and_success)
        else:
            self.error_label.config(text="✗ Невірний пароль!", fg='#f44336')
            self.password_entry.delete(0, tk.END)
            self.shake_password_field()
            self.password_entry.focus_set()

    def shake_password_field(self):
        original_border = self.password_entry.cget('border_color')
        self.password_entry.configure(border_color="#ff0000")
        self.root.after(200, lambda: self.password_entry.configure(border_color=original_border))

    def close_and_success(self):
        """Плавно приховує вікно перед закриттям і викликає callback"""

        def fade_out(opacity):
            if opacity > 0:
                self.root.attributes("-alpha", opacity)
                self.root.after(30, lambda: fade_out(opacity - 0.05))
            else:
                try:
                    self.root.quit()
                    self.root.destroy()
                    if self.success_callback:
                        self.success_callback()
                except tk.TclError:
                    pass

        fade_out(1.0)

    def on_window_close(self):
        try:
            self.root.quit()
            self.root.destroy()
        except tk.TclError:
            pass
        sys.exit(0)

    def show_password_form(self):
        try:
            self.subtitle_label.config(text="Завантаження завершено!")
            self.status_label.config(text="Очікування введення пароля...", fg='#FFA726')

            self.password_frame.pack(pady=(10, 0))

            self.main_frame.update_idletasks()
            self.root.update_idletasks()

            self.password_entry.focus_set()
            self.password_entry.bind('<Return>', lambda e: self.check_password())

            self.root.lift()
            self.root.focus_force()
            self.root.update()
            print("Форма пароля має бути показана")
        except Exception as e:
            print(f"Помилка: {e}")

    def fade_in(self, opacity=0.0):
        """Плавно показує вікно"""
        if opacity < 1.0:
            self.root.attributes("-alpha", opacity)
            self.root.after(30, lambda: self.fade_in(opacity + 0.05))
        else:
            self.root.attributes("-alpha", 1.0)

    def start_loading(self, status_updates=None):
        if status_updates is None:
            status_updates = [
                "Завантаження модулів...",
                "Ініціалізація компонентів...",
                "Перевірка залежностей...",
                "Налаштування інтерфейсу...",
                "Завантаження даних...",
                "Фінальна підготовка..."
            ]

        threading.Thread(
            target=self.loading_animation,
            args=(status_updates,),
            daemon=True
        ).start()

    def loading_animation(self, status_updates):
        steps = 100
        step_duration = self.duration / steps
        status_step = len(status_updates)
        steps_per_status = steps // status_step if status_step > 0 else steps

        current_status_index = 0

        for i in range(steps + 1):
            try:
                if i > 0 and i % steps_per_status == 0 and current_status_index < len(status_updates):
                    self.root.after(0, lambda text=status_updates[current_status_index]: self.status_label.config(text=text))
                    current_status_index += 1

                self.root.after(0, lambda value=i: self.progress.configure(value=value))
                self.root.after(0, lambda percent=i: self.percent_label.config(text=f"{percent}%"))

                remaining_time = self.duration - (i * step_duration)
                time_text = f"Залишилось: {remaining_time:.1f} сек" if remaining_time > 0 else "Завершено!"
                self.root.after(0, lambda text=time_text: self.time_label.config(text=text))

                time.sleep(step_duration)
            except tk.TclError:
                break

        try:
            self.show_password_form()
        except tk.TclError:
            pass

    def close_splash(self):
        try:
            self.root.quit()
            self.root.destroy()
        except tk.TclError:
            pass

    def show(self, success_callback=None):
        self.success_callback = success_callback
        self.root.mainloop()


def show_splash_with_password(duration=4, app_password="qwer", success_callback=None, status_updates=None):
    if status_updates is None:
        status_updates = [
            "Завантаження error_handler...",
            "Завантаження gui_utils...",
            "Завантаження custom_widgets...",
            "Завантаження auth_utils...",
            "Завантаження модулів даних...",
            "Ініціалізація інтерфейсу..."
        ]

    splash = SplashScreen(duration, app_password)
    splash.start_loading(status_updates)
    splash.show(success_callback)


def show_splash_screen(duration):
    show_splash_with_password(duration)


if __name__ == "__main__":
    show_splash_screen(5)
