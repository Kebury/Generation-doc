<<<<<<< HEAD
@echo off
chcp 65001 >nul
echo ═══════════════════════════════════════════════════════════
echo   Создание исполняемого файла "Генератор документов.exe"
echo ═══════════════════════════════════════════════════════════
echo.

echo [1/5] Очистка предыдущих сборок...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "Генератор документов.spec" del "Генератор документов.spec"
if exist "generation_doc.spec" del "generation_doc.spec"
echo.

echo [2/5] Установка зависимостей...
echo Установка PyInstaller и всех необходимых библиотек...
pip install --upgrade pyinstaller
pip install customtkinter pandas python-docx openpyxl pymorphy3 pypdf PyPDF2 pymorphy3-dicts-ru tkinterdnd2 docx2pdf pywin32 Pillow lxml PyMuPDF
pip install winsdk reportlab
echo ✓ Зависимости установлены
echo.

echo [3/5] Сборка приложения...
echo Создание standalone exe с фоновой загрузкой модулей...
echo ВАЖНО: Модули загружаются в фоне для быстрого старта GUI
echo ВАЖНО: OCR использует Windows OCR (встроен в Windows 10+)
pyinstaller --onefile --windowed --noupx ^
    --bootloader-ignore-signals ^
    --name="Генератор документов" ^
    --icon=doc.ico ^
    --hidden-import=os ^
    --hidden-import=re ^
    --hidden-import=io ^
    --hidden-import=sys ^
    --hidden-import=time ^
    --hidden-import=json ^
    --hidden-import=datetime ^
    --hidden-import=calendar ^
    --hidden-import=copy ^
    --hidden-import=traceback ^
    --hidden-import=shutil ^
    --hidden-import=tempfile ^
    --hidden-import=atexit ^
    --hidden-import=platform ^
    --hidden-import=queue ^
    --hidden-import=gc ^
    --hidden-import=subprocess ^
    --hidden-import=winreg ^
    --hidden-import=threading ^
    --hidden-import=multiprocessing ^
    --hidden-import=multiprocessing.spawn ^
    --hidden-import=multiprocessing.pool ^
    --hidden-import=concurrent.futures ^
    --hidden-import=concurrent.futures.process ^
    --hidden-import=concurrent.futures.thread ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --hidden-import=tkinter.filedialog ^
    --hidden-import=tkinter.messagebox ^
    --hidden-import=tkinter.simpledialog ^
    --hidden-import=tkinter.scrolledtext ^
    --hidden-import=customtkinter ^
    --hidden-import=customtkinter.windows ^
    --hidden-import=customtkinter.windows.widgets ^
    --hidden-import=pandas ^
    --hidden-import=pandas.core ^
    --hidden-import=pandas.io.excel ^
    --hidden-import=pandas._libs.tslibs.timedeltas ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.styles ^
    --hidden-import=openpyxl.cell.cell ^
    --hidden-import=docx ^
    --hidden-import=docx.shared ^
    --hidden-import=docx.enum.text ^
    --hidden-import=docx.oxml ^
    --hidden-import=docx.oxml.ns ^
    --hidden-import=docx.oxml.xmlchemy ^
    --hidden-import=docx.document ^
    --hidden-import=lxml ^
    --hidden-import=lxml.etree ^
    --hidden-import=lxml._elementpath ^
    --hidden-import=pymorphy3 ^
    --hidden-import=pymorphy3.analyzer ^
    --hidden-import=pymorphy3.tagset ^
    --hidden-import=pymorphy3.opencorpora_dict ^
    --hidden-import=pymorphy3.units ^
    --hidden-import=pypdf ^
    --hidden-import=PyPDF2 ^
    --hidden-import=tkinterdnd2 ^
    --hidden-import=docx2pdf ^
    --hidden-import=win32com ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    --hidden-import=pywintypes ^
    --hidden-import=PIL ^
    --hidden-import=PIL.Image ^
    --hidden-import=PIL.ImageDraw ^
    --hidden-import=PIL.ImageTk ^
    --hidden-import=PIL.ImageFilter ^
    --hidden-import=PIL.ImageEnhance ^
    --hidden-import=PIL.ImageOps ^
    --hidden-import=fitz ^
    --hidden-import=asyncio ^
    --hidden-import=winsdk ^
    --hidden-import=winsdk.windows.media.ocr ^
    --hidden-import=winsdk.windows.storage.streams ^
    --hidden-import=winsdk.windows.graphics.imaging ^
    --hidden-import=winsdk.windows.globalization ^
    --hidden-import=reportlab ^
    --hidden-import=reportlab.pdfgen ^
    --hidden-import=reportlab.pdfgen.canvas ^
    --hidden-import=reportlab.lib.pagesizes ^
    --hidden-import=reportlab.pdfbase ^
    --hidden-import=reportlab.pdfbase.pdfmetrics ^
    --hidden-import=reportlab.pdfbase.ttfonts ^
    --hidden-import=reportlab.pdfbase.ttfonts.TTFont ^
    --collect-all=pymorphy3_dicts_ru ^
    --collect-submodules=customtkinter ^
    --collect-submodules=multiprocessing ^
    --collect-submodules=pandas ^
    --collect-submodules=docx ^
    --collect-submodules=PIL ^
    --collect-submodules=fitz ^
    --collect-submodules=winsdk ^
    --collect-submodules=reportlab ^
    --exclude-module=matplotlib ^
    --exclude-module=scipy ^
    --exclude-module=numpy.random._examples ^
    generation_doc.py
echo.

echo [4/5] Проверка результата...
if exist "dist\Генератор документов.exe" (
    echo ✓ Файл успешно создан!
) else (
    echo ✗ ОШИБКА: Файл не был создан!
    echo Проверьте логи выше для деталей.
)
echo.

echo [5/5] Готово!
echo.
echo ═══════════════════════════════════════════════════════════
echo   Исполняемый файл: dist\Генератор документов.exe
echo ═══════════════════════════════════════════════════════════
echo.
echo   Возможности OCR:
echo   - Автоматическое распознавание сканированных PDF
echo   - Поддержка кириллицы (русский язык)
echo   - Невидимый текстовый слой для поиска/копирования
echo   - Использует Windows OCR (встроен в Windows 10+)
echo.
pause
=======
@echo off
chcp 65001 >nul
echo ═══════════════════════════════════════════════════════════
echo   Создание исполняемого файла "Генератор документов.exe"
echo ═══════════════════════════════════════════════════════════
echo.

echo [1/5] Очистка предыдущих сборок...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "Генератор документов.spec" del "Генератор документов.spec"
if exist "generation_doc.spec" del "generation_doc.spec"
echo.

echo [2/5] Установка зависимостей...
echo Установка PyInstaller и всех необходимых библиотек...
pip install --upgrade pyinstaller
pip install customtkinter pandas python-docx openpyxl pymorphy3 pypdf PyPDF2 pymorphy3-dicts-ru tkinterdnd2 docx2pdf pywin32 Pillow lxml PyMuPDF
pip install winsdk reportlab
echo ✓ Зависимости установлены
echo.

echo [3/5] Сборка приложения...
echo Создание standalone exe с фоновой загрузкой модулей...
echo ВАЖНО: Модули загружаются в фоне для быстрого старта GUI
echo ВАЖНО: OCR использует Windows OCR (встроен в Windows 10+)
pyinstaller --onefile --windowed --noupx ^
    --bootloader-ignore-signals ^
    --name="Генератор документов" ^
    --icon=doc.ico ^
    --hidden-import=os ^
    --hidden-import=re ^
    --hidden-import=io ^
    --hidden-import=sys ^
    --hidden-import=time ^
    --hidden-import=json ^
    --hidden-import=datetime ^
    --hidden-import=calendar ^
    --hidden-import=copy ^
    --hidden-import=traceback ^
    --hidden-import=shutil ^
    --hidden-import=tempfile ^
    --hidden-import=atexit ^
    --hidden-import=platform ^
    --hidden-import=queue ^
    --hidden-import=gc ^
    --hidden-import=subprocess ^
    --hidden-import=winreg ^
    --hidden-import=threading ^
    --hidden-import=multiprocessing ^
    --hidden-import=multiprocessing.spawn ^
    --hidden-import=multiprocessing.pool ^
    --hidden-import=concurrent.futures ^
    --hidden-import=concurrent.futures.process ^
    --hidden-import=concurrent.futures.thread ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --hidden-import=tkinter.filedialog ^
    --hidden-import=tkinter.messagebox ^
    --hidden-import=tkinter.simpledialog ^
    --hidden-import=tkinter.scrolledtext ^
    --hidden-import=customtkinter ^
    --hidden-import=customtkinter.windows ^
    --hidden-import=customtkinter.windows.widgets ^
    --hidden-import=pandas ^
    --hidden-import=pandas.core ^
    --hidden-import=pandas.io.excel ^
    --hidden-import=pandas._libs.tslibs.timedeltas ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.styles ^
    --hidden-import=openpyxl.cell.cell ^
    --hidden-import=docx ^
    --hidden-import=docx.shared ^
    --hidden-import=docx.enum.text ^
    --hidden-import=docx.oxml ^
    --hidden-import=docx.oxml.ns ^
    --hidden-import=docx.oxml.xmlchemy ^
    --hidden-import=docx.document ^
    --hidden-import=lxml ^
    --hidden-import=lxml.etree ^
    --hidden-import=lxml._elementpath ^
    --hidden-import=pymorphy3 ^
    --hidden-import=pymorphy3.analyzer ^
    --hidden-import=pymorphy3.tagset ^
    --hidden-import=pymorphy3.opencorpora_dict ^
    --hidden-import=pymorphy3.units ^
    --hidden-import=pypdf ^
    --hidden-import=PyPDF2 ^
    --hidden-import=tkinterdnd2 ^
    --hidden-import=docx2pdf ^
    --hidden-import=win32com ^
    --hidden-import=win32com.client ^
    --hidden-import=pythoncom ^
    --hidden-import=pywintypes ^
    --hidden-import=PIL ^
    --hidden-import=PIL.Image ^
    --hidden-import=PIL.ImageDraw ^
    --hidden-import=PIL.ImageTk ^
    --hidden-import=PIL.ImageFilter ^
    --hidden-import=PIL.ImageEnhance ^
    --hidden-import=PIL.ImageOps ^
    --hidden-import=fitz ^
    --hidden-import=asyncio ^
    --hidden-import=winsdk ^
    --hidden-import=winsdk.windows.media.ocr ^
    --hidden-import=winsdk.windows.storage.streams ^
    --hidden-import=winsdk.windows.graphics.imaging ^
    --hidden-import=winsdk.windows.globalization ^
    --hidden-import=reportlab ^
    --hidden-import=reportlab.pdfgen ^
    --hidden-import=reportlab.pdfgen.canvas ^
    --hidden-import=reportlab.lib.pagesizes ^
    --hidden-import=reportlab.pdfbase ^
    --hidden-import=reportlab.pdfbase.pdfmetrics ^
    --hidden-import=reportlab.pdfbase.ttfonts ^
    --hidden-import=reportlab.pdfbase.ttfonts.TTFont ^
    --collect-all=pymorphy3_dicts_ru ^
    --collect-submodules=customtkinter ^
    --collect-submodules=multiprocessing ^
    --collect-submodules=pandas ^
    --collect-submodules=docx ^
    --collect-submodules=PIL ^
    --collect-submodules=fitz ^
    --collect-submodules=winsdk ^
    --collect-submodules=reportlab ^
    --exclude-module=matplotlib ^
    --exclude-module=scipy ^
    --exclude-module=numpy.random._examples ^
    generation_doc.py
echo.

echo [4/5] Проверка результата...
if exist "dist\Генератор документов.exe" (
    echo ✓ Файл успешно создан!
) else (
    echo ✗ ОШИБКА: Файл не был создан!
    echo Проверьте логи выше для деталей.
)
echo.

echo [5/5] Готово!
echo.
echo ═══════════════════════════════════════════════════════════
echo   Исполняемый файл: dist\Генератор документов.exe
echo ═══════════════════════════════════════════════════════════
echo.
echo   Возможности OCR:
echo   - Автоматическое распознавание сканированных PDF
echo   - Поддержка кириллицы (русский язык)
echo   - Невидимый текстовый слой для поиска/копирования
echo   - Использует Windows OCR (встроен в Windows 10+)
echo.
pause
>>>>>>> 93b32c39be899ae26df05cbca3677821b1448be0
