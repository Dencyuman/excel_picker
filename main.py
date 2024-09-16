import pystray
from PIL import Image
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import queue
import os
import sys  # sysモジュールをインポート

def resource_path(relative_path):
    """リソースファイルのパスを取得する"""
    try:
        # PyInstallerでパッケージ化された場合
        base_path = sys._MEIPASS
    except Exception:
        # 通常の実行環境
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def main():
    print("プログラムを開始します。")
    # tkinterのルートウィンドウを作成
    root = tk.Tk()
    root.withdraw()  # ルートウィンドウを非表示にする
    print("tkinterのルートウィンドウを作成しました。")

    # スレッド間通信のためのキューを作成
    q = queue.Queue()
    print("キューを作成しました。")

    # メニューアイテムが選択されたときの処理
    def on_clicked(icon, item):
        print("メニューアイテムが選択されました。")
        # キューにメッセージを追加
        q.put('clicked')

    # アイコンを右クリックしてメニューから終了を選択したときの処理
    def on_exit(icon, item):
        print("プログラムを終了します。")
        icon.stop()
        root.quit()

    def process_queue():
        try:
            while True:
                msg = q.get_nowait()
                print(f"キューからメッセージを取得: {msg}")
                if msg == 'clicked':
                    select_files()
        except queue.Empty:
            pass
        # 次のチェックをスケジュール
        root.after(100, process_queue)

    def select_files():
        print("ファイルダイアログを表示します。")
        # Excelファイルを選択するファイルダイアログを開く（複数選択可能）
        file_paths = filedialog.askopenfilenames(
            filetypes=[("Excel files", "*.xlsx;*.xls")],
            title="Excelファイルを選択してください"
        )
        if file_paths:
            # 選択されたファイル名のリストを取得
            file_names = [os.path.basename(path) for path in file_paths]
            print(f"ファイルが選択されました: {file_names}")

            # 確認ダイアログを表示
            def on_ok():
                print("送信しました。")
                # 確認ウィンドウを閉じる
                confirm_window.destroy()
                # 結果ウィンドウを表示
                result_window = tk.Toplevel()
                result_window.title("結果")
                result_window.resizable(False, False)
                # メッセージラベル
                result_label = tk.Label(result_window, text="送信しました。", pady=10)
                result_label.pack(padx=20, pady=10)
                # 閉じるボタン
                close_button = tk.Button(result_window, text="閉じる", width=10, command=result_window.destroy)
                close_button.pack(pady=10)
                # ウィンドウの幅を調整
                result_window.update_idletasks()
                width = max(300, result_window.winfo_width())
                height = result_window.winfo_height()
                result_window.geometry(f"{width}x{height}")
                # ウィンドウの位置を中央に
                center_window(result_window)

            def on_cancel():
                print("キャンセルしました。")
                # 確認ウィンドウを閉じる
                confirm_window.destroy()
                # 結果ウィンドウを表示
                result_window = tk.Toplevel()
                result_window.title("結果")
                result_window.resizable(False, False)
                # メッセージラベル
                result_label = tk.Label(result_window, text="キャンセルしました。", pady=10)
                result_label.pack(padx=20, pady=10)
                # 閉じるボタン
                close_button = tk.Button(result_window, text="閉じる", width=10, command=result_window.destroy)
                close_button.pack(pady=10)
                # ウィンドウの幅を調整
                result_window.update_idletasks()
                width = max(300, result_window.winfo_width())
                height = result_window.winfo_height()
                result_window.geometry(f"{width}x{height}")
                # ウィンドウの位置を中央に
                center_window(result_window)

            # ウィンドウを作成
            confirm_window = tk.Toplevel()
            confirm_window.title("確認")
            confirm_window.resizable(False, False)

            # メインフレーム
            main_frame = tk.Frame(confirm_window, padx=20, pady=20)
            main_frame.pack()

            # メッセージラベル
            message = "以下のファイルを送信しますか？"
            message_label = tk.Label(main_frame, text=message)
            message_label.pack(pady=(0, 10))

            # ファイル名リストを表示
            files_frame = tk.Frame(main_frame)
            files_frame.pack()

            for fname in file_names:
                fname_label = tk.Label(files_frame, text=fname, anchor='w')
                fname_label.pack(fill='x')

            # ボタンフレーム
            button_frame = tk.Frame(main_frame)
            button_frame.pack(pady=20)

            # OKボタン
            ok_button = tk.Button(button_frame, text="OK", width=10, command=on_ok)
            ok_button.pack(side=tk.LEFT, padx=10)

            # キャンセルボタン
            cancel_button = tk.Button(button_frame, text="キャンセル", width=10, command=on_cancel)
            cancel_button.pack(side=tk.LEFT, padx=10)

            # ウィンドウの幅を調整
            confirm_window.update_idletasks()
            width = max(400, confirm_window.winfo_width())
            height = confirm_window.winfo_height()
            confirm_window.geometry(f"{width}x{height}")

            # ウィンドウの位置を中央に
            center_window(confirm_window)

        else:
            print("ファイルが選択されませんでした。")
            # ファイルが選択されなかった場合の処理
            pass

    def center_window(win):
        win.update_idletasks()
        width = win.winfo_width()
        height = win.winfo_height()
        x = (win.winfo_screenwidth() // 2) - (width // 2)
        y = (win.winfo_screenheight() // 2) - (height // 2)
        win.geometry(f'{width}x{height}+{x}+{y}')

    # キューの処理を開始
    print("キューの処理を開始します。")
    root.after(100, process_queue)

    # pystrayのアイコンを実行する関数
    def run_icon():
        print("トレイアイコンを設定します。")
        # アイコン画像を読み込む
        icon_path = resource_path("icon.png")
        image = Image.open(icon_path)
        icon = pystray.Icon("ExcelFileSelector", image, "Excelファイル選択",
                            menu=pystray.Menu(
                                pystray.MenuItem('ファイルを選択', on_clicked),
                                pystray.MenuItem('終了', on_exit)
                            ))
        print("トレイアイコンを実行します。")
        icon.run()
        print("トレイアイコンの実行を終了しました。")

    # pystrayのアイコンを別スレッドで開始
    print("トレイアイコンのスレッドを開始します。")
    icon_thread = threading.Thread(target=run_icon, daemon=True)
    icon_thread.start()

    # tkinterのメインループを開始
    print("tkinterのメインループを開始します。")
    root.mainloop()
    print("プログラムを終了します。")

if __name__ == "__main__":
    # 起動メッセージを表示
    print("起動しました。システムトレイから実行可能です。")
    startup_root = tk.Tk()
    startup_root.withdraw()
    messagebox.showinfo("起動", "起動しました。システムトレイから実行可能です。")
    startup_root.destroy()
    main()
