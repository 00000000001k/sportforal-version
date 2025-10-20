# updater.py
"""
Спрощений автооновлювач для Python програм
Працює тільки з .exe файлами з GitHub релізів
"""

import os
import sys
import tempfile
import subprocess
import requests
from packaging import version
from pathlib import Path
import logging


class AutoUpdater:
    """Простий клас для автооновлення .exe програм з GitHub"""

    def __init__(self, repo_owner: str, repo_name: str, current_version: str):
        self.repo_owner = repo_owner
        self.repo_name = repo_name
        self.current_version = current_version.lstrip('v')
        self.github_api_url = f"https://api.github.com/repos/{repo_owner}/{repo_name}"

        # Налаштування логування
        logging.basicConfig(level=logging.INFO)
        self.logger = logging.getLogger(__name__)

        # Перевіряємо, чи це .exe програма
        self.is_exe = getattr(sys, 'frozen', False)
        if self.is_exe:
            self.current_exe_path = sys.executable
        else:
            self.current_exe_path = None

        self.logger.info(f"Ініціалізовано оновлювач. Версія: {self.current_version}, exe: {self.is_exe}")

    def check_for_updates(self) -> dict:
        """Перевіряє наявність нових версій на GitHub"""
        try:
            self.logger.info("Перевірка оновлень на GitHub...")

            url = f"{self.github_api_url}/releases/latest"
            response = requests.get(url, timeout=15)
            response.raise_for_status()

            release_data = response.json()
            latest_version = release_data['tag_name'].lstrip('v')

            self.logger.info(f"Поточна версія: {self.current_version}, остання: {latest_version}")

            # Перевіряємо чи є нова версія
            has_update = version.parse(latest_version) > version.parse(self.current_version)

            # Шукаємо .exe файл в assets
            exe_url = None
            for asset in release_data.get('assets', []):
                if asset['name'].endswith('.exe'):
                    exe_url = asset['browser_download_url']
                    break

            return {
                'has_update': has_update,
                'version': latest_version,
                'current_version': self.current_version,
                'release_notes': release_data.get('body', 'Немає опису'),
                'exe_download_url': exe_url,
                'exe_found': exe_url is not None
            }

        except requests.exceptions.RequestException as e:
            self.logger.error(f"Помилка мережі: {e}")
            return {'error': f'Помилка з\'єднання: {str(e)}'}
        except Exception as e:
            self.logger.error(f"Помилка перевірки: {e}")
            return {'error': f'Помилка: {str(e)}'}

    def download_update(self, download_url: str, progress_callback=None) -> str:
        """Завантажує новий .exe файл"""
        try:
            self.logger.info(f"Завантаження з: {download_url}")

            response = requests.get(download_url, stream=True, timeout=30)
            response.raise_for_status()

            # Створюємо тимчасовий файл
            temp_dir = tempfile.gettempdir()
            temp_file = os.path.join(temp_dir, "SportForAll_new.exe")

            # Видаляємо старий файл якщо є
            if os.path.exists(temp_file):
                os.remove(temp_file)

            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0

            with open(temp_file, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)

                        if progress_callback and total_size > 0:
                            progress = (downloaded / total_size) * 100
                            progress_callback(progress)

            self.logger.info(f"Файл завантажено: {temp_file}")
            return temp_file

        except Exception as e:
            self.logger.error(f"Помилка завантаження: {e}")
            raise Exception(f"Не вдалося завантажити: {str(e)}")

    def install_update(self, new_exe_path: str) -> bool:
        """Встановлює оновлення через .bat файл"""
        try:
            if not self.is_exe:
                raise Exception("Оновлення працює тільки для .exe версії програми")

            if not os.path.exists(new_exe_path):
                raise Exception("Завантажений файл не знайдено")

            # Створюємо .bat файл для заміни
            bat_content = f'''@echo off
            chcp 65001 >nul
            echo Оновлення SportForAll...
            timeout /t 3 /nobreak >nul
            
            echo Заміна файлу програми...
            copy /y "{new_exe_path}" "{self.current_exe_path}"
            if errorlevel 1 (
                echo ПОМИЛКА: Не вдалося замінити файл програми
                pause
                exit /b 1
            )
            
            echo Очищення тимчасових файлів...
            del "{new_exe_path}"
            
            echo Запуск оновленої програми...
            start "" "{self.current_exe_path}"
            
            echo Оновлення завершено!
            del "%~f0"
'''

            bat_file = os.path.join(tempfile.gettempdir(), 'sportforall_update.bat')

            # Створюємо .bat файл в кодуванні UTF-8
            with open(bat_file, 'w', encoding='utf-8') as f:
                f.write(bat_content)

            self.logger.info(f"Створено .bat файл: {bat_file}")

            # Запускаємо .bat файл
            subprocess.Popen(
                ['cmd', '/c', bat_file],
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )

            self.logger.info("Запущено процес оновлення")
            return True

        except Exception as e:
            self.logger.error(f"Помилка встановлення: {e}")
            raise Exception(f"Не вдалося встановити: {str(e)}")

    def perform_full_update(self, progress_callback=None) -> dict:
        """Виконує повний цикл оновлення"""
        try:
            # Крок 1: Перевірка оновлень
            if progress_callback:
                progress_callback(10, "Перевірка оновлень на GitHub...")

            update_info = self.check_for_updates()

            if 'error' in update_info:
                return {'success': False, 'error': update_info['error']}

            if not update_info['has_update']:
                return {
                    'success': False,
                    'no_update': True,
                    'message': f"У вас остання версія: {update_info['current_version']}"
                }

            if not update_info['exe_found']:
                return {
                    'success': False,
                    'error': 'У релізі не знайдено .exe файл для завантаження'
                }

            # Крок 2: Завантаження
            if progress_callback:
                progress_callback(20, f"Завантаження версії {update_info['version']}...")

            def download_progress(percent):
                if progress_callback:
                    # Масштабуємо прогрес завантаження на 20-80%
                    scaled_progress = 20 + (percent * 0.6)
                    progress_callback(scaled_progress, f"Завантаження: {percent:.1f}%")

            new_exe_path = self.download_update(
                update_info['exe_download_url'],
                download_progress
            )

            # Крок 3: Встановлення
            if progress_callback:
                progress_callback(90, "Підготовка до встановлення...")

            self.install_update(new_exe_path)

            if progress_callback:
                progress_callback(100, "Оновлення завершено! Програма перезапуститься...")

            return {
                'success': True,
                'version': update_info['version'],
                'message': f"Оновлення до версії {update_info['version']} встановлено"
            }

        except Exception as e:
            self.logger.error(f"Помилка оновлення: {e}")
            return {'success': False, 'error': str(e)}


# Проста функція для перевірки оновлень
def check_updates(repo_owner: str, repo_name: str, current_version: str) -> dict:
    """Простий спосіб перевірити оновлення"""
    updater = AutoUpdater(repo_owner, repo_name, current_version)
    result = updater.check_for_updates()

    if 'error' in result:
        return None

    return {
        'newer': result['has_update'],
        'version': result['version'],
        'release_notes': result['release_notes']
    }


# Приклад використання
if __name__ == "__main__":
    def progress(percent, status):
        print(f"\r{status} [{percent:.1f}%]", end="")


    updater = AutoUpdater("KostyaBulavko", "sportfprall", "2.16.3")
    result = updater.perform_full_update(progress)

    if result['success']:
        print(f"\n✅ {result['message']}")
    else:
        print(f"\n❌ Помилка: {result.get('error', 'Невідома помилка')}")