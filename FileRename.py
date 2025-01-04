import os
import re
import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QPushButton, QLabel, QTableWidget, QTableWidgetItem,
    QWidget, QHBoxLayout, QMessageBox, QComboBox, QStatusBar, QCheckBox, QSpacerItem, QSizePolicy,
    QMenu, QAction
)
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QColor, QBrush, QIcon
from PyQt5.QtMultimedia import QSound
import subfunc
from natsort import natsorted

# アプリ名称
WINDOW_TITLE = "File Rename Tool"
# 設定ファイル
SETTINGS_FILE = "FileRename_settings.json"
# ログファイル
LOGS_FILE = "FileRename.log"
# 設定ファイルのキー名
GEOMETRY_X = "geometry-x"
GEOMETRY_Y = "geometry-y"
GEOMETRY_W = "geometry-w"
GEOMETRY_H = "geometry-h"
HISTORY_BEFORE = "history-before"
HISTORY_AFTER = "history-after"
OPT_USEREGULAR = "opt-useregular"
OPT_IGNORECASE = "opt-ignorecase"
SOUND_NG = "sound-ng"
SOUND_OK = "sound-ok"

class FileRename(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(WINDOW_TITLE)
        self.setGeometry(100, 100, 848, 480)

        self.history_before = []
        self.history_after = []
        self.opt_useregular = False
        self.opt_ignorecase = False
        self.folder_path = None
        self.matchfilenum = 0
        self.soundok = "ok.wav"
        self.soundng = "ng.wav"
        self.pydir = os.path.dirname(os.path.abspath(__file__))

        self.init_ui()

    # 終了時処理
    def closeEvent(self, event):
        self.save_settings()
        super().closeEvent(event)

    #ドラッグ＆ドロップエンター時処理
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    #ドラッグ＆ドロップ時処理
    def dropEvent(self, event: QDropEvent):
        self.clear_list()

        urls = event.mimeData().urls()
        if not urls:
            #QMessageBox.warning(self, "Warning", "No valid files or folders dropped.")
            self.show_err_message(f"不正なファイルかフォルダがドロップされました")
            return

        paths = [url.toLocalFile() for url in urls]
        if all(os.path.isfile(path) for path in paths):
            # All are files
            parent_dirs = {os.path.dirname(path) for path in paths}
            if len(parent_dirs) > 1:
                #QMessageBox.warning(self, "Warning", "All files must be in the same folder.")
                self.show_err_message(f"全てのファイルは同じフォルダ内にある必要があります")
                return

            self.folder_path = parent_dirs.pop()
            self.files = paths
            self.show_message(f"{len(self.files)}ファイルがドロップされました")
        elif len(paths) == 1 and os.path.isdir(paths[0]):
            # Single folder
            self.folder_path = paths[0]
            self.files.clear()
            self.add_files_from_directory(self.folder_path)
            self.show_message(f"{len(self.files)}フォルダがドロップされました")
        elif all(os.path.isdir(path) for path in paths):
            # All are files
            parent_dirs = {os.path.dirname(path) for path in paths}
            if len(parent_dirs) > 1:
                #QMessageBox.warning(self, "Warning", "All files must be in the same folder.")
                self.show_err_message(f"全てのフォルダは同じ親フォルダ内にある必要があります")
                return

            self.folder_path = parent_dirs.pop()
            self.files = paths
            self.show_message(f"{len(self.files)}フォルダがドロップされました")
        else:
            #QMessageBox.warning(self, "Warning", "Invalid mix of files and folders.")
            self.show_err_message(f"ファイルとフォルダが同時にドロップされました")
            return

        # Windowsのエクスプローラーのようなナチュラルソート
        self.files = natsorted(self.files)
        #self.lanbel_folder.setText(self.folder_path)
        self.set_folder_label(self.folder_path)
        #self.status_bar.showMessage(f"Selected folder: {self.folder_path}")
        self.update_file_table()
        if len(self.files) == 0:
            self.show_err_message(f"空のフォルダがドロップされました")

    # 設定値更新時処理
    def update_regex_check(self):
        self.opt_useregular = self.regex_checkbox.isChecked()
    def update_ignorecase_check(self):
        self.opt_ignorecase = self.ignorecase_checkbox.isChecked()

    # 右クリックメニュー登録処理
    def contextMenuEvent(self, event):
        menu = QMenu(self)

        action1 = QAction("Historyのクリア", self)
        menu.addAction(action1)
        action1.triggered.connect(self.clear_history)
        menu.exec(event.globalPos())

    #----------------------------------------
    #- 処理関数
    #----------------------------------------
    def init_ui(self):
        if not os.path.exists(SETTINGS_FILE):
            self.createSettingFile()
        self.load_settings()

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()

        # Top Layout for Inputs and Options
        top_layout_input = QHBoxLayout()
        top_layout_option = QHBoxLayout()

        label = QLabel("検索:")
        label.setFixedWidth(48)
        top_layout_input.addWidget(label)
        self.replace_before_combo = QComboBox()
        self.replace_before_combo.setEditable(True)
        self.replace_before_combo.setCompleter(None) #サジェスト機能の無効化
        self.replace_before_combo.setMinimumWidth(200)
        #self.replace_before_combo.setMaximumWidth(400)
        top_layout_input.addWidget(self.replace_before_combo)
        label = QLabel("置換:")
        label.setFixedWidth(48)
        top_layout_input.addWidget(label)
        self.replace_after_combo = QComboBox()
        self.replace_after_combo.setEditable(True)
        self.replace_after_combo.setCompleter(None) #サジェスト機能の無効化
        self.replace_after_combo.setMinimumWidth(200)
        #self.replace_after_combo.setMaximumWidth(400)
        top_layout_input.addWidget(self.replace_after_combo)
        #top_layout_input.addSpacerItem(QSpacerItem(0, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        layout.addLayout(top_layout_input)

        self.update_history_combo()

        self.regex_checkbox = QCheckBox("正規表現")
        top_layout_option.addWidget(self.regex_checkbox)
        self.ignorecase_checkbox = QCheckBox("大文字小文字無視")
        top_layout_option.addWidget(self.ignorecase_checkbox)
        top_layout_option.addSpacerItem(QSpacerItem(0, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        layout.addLayout(top_layout_option)

        self.regex_checkbox.stateChanged.connect(self.update_regex_check)
        self.ignorecase_checkbox.stateChanged.connect(self.update_ignorecase_check)
        self.regex_checkbox.setChecked(self.opt_useregular)
        self.ignorecase_checkbox.setChecked(self.opt_ignorecase)

        # File Table
        self.file_table = QTableWidget()
        self.file_table.setColumnCount(2)
        self.file_table.setHorizontalHeaderLabels(["現在の名前", "置換後"])
        self.file_table.horizontalHeader().setStretchLastSection(True)
        self.file_table.horizontalHeader().setSectionResizeMode(0, 1)  # Adjustable columns
        self.file_table.horizontalHeader().setSectionResizeMode(1, 1)
        self.file_table.verticalHeader().setDefaultSectionSize(18)  # Reduced row height
        layout.addWidget(self.file_table)

        # Buttons
        layout_buttons = QHBoxLayout()
        layout_buttons.addWidget(QLabel("フォルダ:"))
        self.lanbel_folder = QLabel()
        self.set_folder_label()
        layout_buttons.addWidget(self.lanbel_folder)
        layout_buttons.addSpacerItem(QSpacerItem(0, 40, QSizePolicy.Expanding, QSizePolicy.Minimum))
        self.test_seq_button = QPushButton("連番テスト")
        self.test_seq_button.setFixedHeight(40)
        self.test_seq_button.setStyleSheet(
            """
            QPushButton {
                background-color: #AACCFF;  /* 背景色 */
                color: black;  /* 文字色 */
            }
            """
        )
        self.test_seq_button.setToolTip("このボタンではファイルを001からの連番への変換テストを行います")
        self.test_seq_button.clicked.connect(self.test_seq)
        layout_buttons.addWidget(self.test_seq_button)
        self.seq_button = QPushButton("連番変換")
        self.seq_button.setFixedHeight(40)
        self.seq_button.setStyleSheet(
            """
            QPushButton {
                background-color: #FFAACC;  /* 背景色 */
                color: black;  /* 文字色 */
            }
            """
        )
        self.seq_button.setToolTip("このボタンではファイルを001からの連番への変換を行います")
        self.seq_button.clicked.connect(self.seq_files)
        layout_buttons.addWidget(self.seq_button)
        layout_buttons.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Fixed, QSizePolicy.Minimum))

        self.test_button = QPushButton("置換テスト")
        self.test_button.setFixedHeight(40)
        self.test_button.setStyleSheet(
            """
            QPushButton {
                background-color: #22DD00;  /* 背景色 */
                color: black;  /* 文字色 */
            }
            """
        )
        self.test_button.clicked.connect(self.test_rename)
        layout_buttons.addWidget(self.test_button)
        self.rename_button = QPushButton("置換")
        self.rename_button.setFixedHeight(40)
        self.rename_button.setStyleSheet(
            """
            QPushButton {
                background-color: #DD2200;  /* 背景色 */
                color: black;  /* 文字色 */
            }
            """
        )
        self.rename_button.clicked.connect(self.rename_files)
        layout_buttons.addWidget(self.rename_button)
        self.clear_button = QPushButton("クリア")
        self.clear_button.setFixedHeight(40)
        self.clear_button.clicked.connect(self.clear_func)
        layout_buttons.addWidget(self.clear_button)
        layout.addLayout(layout_buttons)

        # Status Bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        central_widget.setLayout(layout)

        self.setWindowIcon(QIcon("res/FileRename.ico"))
        self.setAcceptDrops(True)
        self.files = []

    #設定ファイルの初期値作成
    def createSettingFile(self):
        self.save_settings()

    #設定ファイルのロード
    def load_settings(self):
        geox = subfunc.read_value_from_config(SETTINGS_FILE, GEOMETRY_X)
        geoy = subfunc.read_value_from_config(SETTINGS_FILE, GEOMETRY_Y)
        geow = subfunc.read_value_from_config(SETTINGS_FILE, GEOMETRY_W)
        geoh = subfunc.read_value_from_config(SETTINGS_FILE, GEOMETRY_H)
        if any(val is None for val in [geox, geoy, geow, geoh]):
            self.setGeometry(100, 100, 848, 480)    #位置とサイズ
        else:
            self.setGeometry(geox, geoy, geow, geoh)    #位置とサイズ

        self.history_before = subfunc.read_list_from_config(SETTINGS_FILE, HISTORY_BEFORE)
        if not self.history_before:
            self.history_before = []
        self.history_after = subfunc.read_list_from_config(SETTINGS_FILE, HISTORY_AFTER)
        if not self.history_after:
            self.history_after = []

        self.opt_useregular = subfunc.read_value_from_config(SETTINGS_FILE, OPT_USEREGULAR)
        if self.opt_useregular is None:
            self.opt_useregular = False
        self.opt_ignorecase = subfunc.read_value_from_config(SETTINGS_FILE, OPT_IGNORECASE)
        if self.opt_ignorecase is None:
            self.opt_ignorecase = False

        self.soundok = subfunc.read_value_from_config(SETTINGS_FILE, SOUND_OK)
        if self.soundok is None:
            self.soundok = "ok.wav"
        self.soundng = subfunc.read_value_from_config(SETTINGS_FILE, SOUND_NG)
        if self.soundng is None:
            self.soundng = "ng.wav"

    #設定ファイルのセーブ
    def save_settings(self):
        subfunc.write_value_to_config(SETTINGS_FILE, GEOMETRY_X, self.geometry().x())
        subfunc.write_value_to_config(SETTINGS_FILE, GEOMETRY_Y, self.geometry().y())
        subfunc.write_value_to_config(SETTINGS_FILE, GEOMETRY_W, self.geometry().width())
        subfunc.write_value_to_config(SETTINGS_FILE, GEOMETRY_H, self.geometry().height())
        subfunc.write_list_from_config(SETTINGS_FILE, HISTORY_BEFORE, self.history_before)
        subfunc.write_list_from_config(SETTINGS_FILE, HISTORY_AFTER, self.history_after)
        subfunc.write_value_to_config(SETTINGS_FILE, OPT_USEREGULAR, self.opt_useregular)
        subfunc.write_value_to_config(SETTINGS_FILE, OPT_IGNORECASE, self.opt_ignorecase)
        subfunc.write_value_to_config(SETTINGS_FILE, SOUND_OK, self.soundok)
        subfunc.write_value_to_config(SETTINGS_FILE, SOUND_NG, self.soundng)

    # フォルダのドロップ
    def add_files_from_directory(self, directory):
        for file in os.listdir(directory):
            file_path = os.path.join(directory, file)
            self.files.append(file_path)
    # テーブルの更新
    def update_file_table(self):
        self.file_table.setRowCount(len(self.files))
        for row, file in enumerate(self.files):
            self.file_table.setItem(row, 0, QTableWidgetItem(os.path.basename(file)))

    # 連番テストボタン処理
    def test_seq(self):
        self.matchfilenum = -1

        if len(self.files) == 0:
            self.show_err_message(f"ファイルが登録されていません")
            return

        # ファイル名の桁数は最小で3桁、ファイル数がそれ以上ならそれに合わせて0パディング
        padding = len(str(self.matchfilenum))
        if padding < 3:
            padding = 3
        self.matchfilenum = 0
        for row, file in enumerate(self.files):
            basename = os.path.basename(file)
            self.file_table.setItem(row, 0, QTableWidgetItem(basename))
            fn, ext = os.path.splitext(basename)
            if fn != "" and ext != "":
                self.matchfilenum += 1
                new_name = str(self.matchfilenum).zfill(padding) + ext
                new_item = QTableWidgetItem(new_name)
                new_item.setForeground(QBrush(QColor("red")))
                self.file_table.setItem(row, 1, new_item)
            else:
                new_name = basename
                new_item = QTableWidgetItem(new_name)
                self.file_table.setItem(row, 1, new_item)

        if self.matchfilenum == 0:
            self.show_err_message(f"連番の対象ファイルが0件です")
            return

        self.show_message(f"テスト結果 : {self.file_table.rowCount()}ファイルのうち、{self.matchfilenum}ファイルが対象です")
        return

    # 連番変換ボタン処理
    def seq_files(self):
        self.test_seq()
        if self.matchfilenum <= 0:
            # すでにテストでエラー状況を表示済み
            return

        response = QMessageBox.question(self, "確認", "本当に連番処理を実施してよいですか？", QMessageBox.Yes | QMessageBox.No)
        if response != QMessageBox.Yes:
            return
        # ファイル名の桁数は最小で3桁、ファイル数がそれ以上ならそれに合わせて0パディング
        padding = len(str(self.matchfilenum))
        if padding < 3:
            padding = 3
        new_names = {}
        self.matchfilenum = 0
        for row, file in enumerate(self.files):
            basename = os.path.basename(file)
            self.file_table.setItem(row, 0, QTableWidgetItem(basename))
            fn, ext = os.path.splitext(basename)
            if fn != "" and ext != "":
                self.matchfilenum += 1
                new_name = str(self.matchfilenum).zfill(padding) + ext
            else:
                new_name = basename

            new_path = os.path.join(os.path.dirname(file), new_name)
            new_names[file] = new_path

        for old_path, new_path in new_names.items():
            os.rename(old_path, new_path)

        self.files = list(new_names.values())
        self.update_file_table()
        #QMessageBox.information(self, "Success", "Files renamed successfully.")
        self.show_message(f"連番結果 : {self.file_table.rowCount()}ファイルのうち、{self.matchfilenum}ファイルを連番にしました")
        self.play_wave(self.soundok)

    # テストボタン処理
    def test_rename(self):
        self.matchfilenum = -1

        if len(self.files) == 0:
            self.show_err_message(f"ファイルが登録されていません")
            return

        before_pattern = self.replace_before_combo.currentText()
        after_text = self.replace_after_combo.currentText()
        if not before_pattern:
            #QMessageBox.warning(self, "Warning", "Please specify a pattern to replace.")
            self.show_err_message(f"検索文字列が指定されていません")
            return

        flags = re.IGNORECASE if self.opt_ignorecase else 0
        try:
            re.compile(before_pattern, flags) if self.opt_useregular else None
        except re.error:
            #QMessageBox.critical(self, "Error", "Invalid regular expression.")
            self.show_err_message(f"正規表現の記載方法に誤りがあります")
            return

        self.file_table.setRowCount(len(self.files))
        conflicts = set()
        new_names = []

        for file in self.files:
            basename = os.path.basename(file)
            if self.opt_useregular:
                new_name = re.sub(before_pattern, after_text, basename, flags=flags)
            else:
                new_name = re.sub(re.escape(before_pattern), after_text, basename, flags=flags)
            new_path = os.path.join(os.path.dirname(file), new_name)

            if new_path in new_names:
                conflicts.add(new_path)

            new_names.append(new_path)

        if conflicts:
            #QMessageBox.warning(self, "Warning", "Filename conflicts detected. Aborting test.")
            self.show_err_message(f"ファイル名の競合があります")
            return

        self.matchfilenum = 0
        for row, file in enumerate(self.files):
            basename = os.path.basename(file)
            if self.opt_useregular:
                new_name = re.sub(before_pattern, after_text, basename, flags=flags)
            else:
                new_name = re.sub(re.escape(before_pattern), after_text, basename, flags=flags)

            self.file_table.setItem(row, 0, QTableWidgetItem(basename))

            new_item = QTableWidgetItem(new_name)

            match = re.search(before_pattern, basename, flags) if self.opt_useregular else re.search(re.escape(before_pattern), basename, flags)
            if match:
                new_item.setForeground(QBrush(QColor("red")))
                self.matchfilenum += 1

            self.file_table.setItem(row, 1, new_item)

        if self.matchfilenum == 0:
            self.show_err_message(f"検索文字列に一致するファイルがありません")
            return

        self.show_message(f"テスト結果 : {self.file_table.rowCount()}ファイルのうち、{self.matchfilenum}ファイルが対象です")
        return

    # 変換ボタン処理
    def rename_files(self):
        before_pattern = self.replace_before_combo.currentText()
        after_text = self.replace_after_combo.currentText()
        self.test_rename()
        if self.matchfilenum <= 0:
            # すでにテストでエラー状況を表示済み
            return

        response = QMessageBox.question(self, "確認", "本当に置換処理を実施してよいですか？", QMessageBox.Yes | QMessageBox.No)
        if response != QMessageBox.Yes:
            return

        flags = re.IGNORECASE if self.opt_ignorecase else 0

        try:
            re.compile(before_pattern, flags) if self.opt_useregular else None
        except re.error:
            #QMessageBox.critical(self, "Error", "Invalid regular expression.")
            self.show_err_message(f"正規表現の記載方法に誤りがあります")
            return

        if before_pattern not in self.history_before:
            self.history_before.insert(0, before_pattern)
        if after_text not in self.history_after:
            self.history_after.insert(0, after_text)
        self.update_history_combo()

        conflicts = set()
        new_names = {}

        for file in self.files:
            dirname, basename = os.path.split(file)
            if self.opt_useregular:
                new_name = re.sub(before_pattern, after_text, basename, flags=flags)
            else:
                new_name = re.sub(re.escape(before_pattern), after_text, basename, flags=flags)

            new_path = os.path.join(dirname, new_name)

            if new_path in new_names.values():
                conflicts.add(new_path)

            new_names[file] = new_path

        if conflicts:
            #QMessageBox.warning(self, "Warning", "Filename conflicts detected. Aborting rename.")
            self.show_err_message(f"ファイル名の競合があります")
            return

        for old_path, new_path in new_names.items():
            if not os.path.exists(new_path):
                os.rename(old_path, new_path)

        self.files = list(new_names.values())
        self.update_file_table()
        #QMessageBox.information(self, "Success", "Files renamed successfully.")
        self.show_message(f"置換結果 : {self.file_table.rowCount()}ファイルのうち、{self.matchfilenum}ファイルを置換しました")
        self.play_wave(self.soundok)

    # クリアボタン処理
    def clear_func(self):
        self.clear_list()
        self.clear_word()
        self.show_message(f"クリアしました")

    def clear_list(self):
        #self.file_table.clear()             # ヘッダーも含めてクリア
        #self.file_table.clearContents()     # アイテムの中身をクリア（リストの件数やヘッダーはそのまま）
        self.file_table.setRowCount(0)      # リストの0件に（ヘッダーはそのまま）
        self.set_folder_label()

    def clear_word(self):
        self.replace_before_combo.setCurrentText("")
        self.replace_after_combo.setCurrentText("")

    # フォルダー名設定
    def set_folder_label(self, str=None):
        if str:
            self.lanbel_folder.setText(str)
        else:
            self.lanbel_folder.setText(" --- ")

    # 変換前後のコンボボックスの更新
    def update_history_combo(self):
        self.history_before = self.history_before[:20]
        self.history_after = self.history_after[:20]
        text = self.replace_before_combo.currentText()
        self.replace_before_combo.clear()
        self.replace_before_combo.addItems(self.history_before)
        self.replace_before_combo.setCurrentText(text)

        text = self.replace_after_combo.currentText()
        self.replace_after_combo.clear()
        self.replace_after_combo.addItems(self.history_after)
        self.replace_after_combo.setCurrentText(text)

    # 右クリックメニュー - ログファイルを開く
    def clear_history(self):
        self.history_before = []
        self.history_after = []
        self.update_history_combo()

    def play_wave(self, file_name):
        file_path = os.path.join(self.pydir, file_name)
        if not os.path.isfile(file_path): return
        sound = QSound(file_path)
        sound.play()
        while sound.isFinished() is False:
            app.processEvents()

    def show_err_message(self, mes):
        self.status_bar.showMessage(f"エラー : {mes}")
        self.play_wave(self.soundng)

    def show_message(self, mes):
        self.status_bar.showMessage(f"{mes}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileRename()
    window.show()
    sys.exit(app.exec_())
