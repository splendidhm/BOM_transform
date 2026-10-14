"""PyQt5 desktop queue for sequential BOM conversion."""
from pathlib import Path
import os
import sys
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'vendor'))
import csv
import json
import traceback
from datetime import datetime
from PyQt5.QtCore import Qt,QThread,pyqtSignal,QUrl,QTimer,QCoreApplication
from PyQt5.QtGui import QFont,QDesktopServices
from PyQt5.QtWidgets import (QApplication,QMainWindow,QWidget,QVBoxLayout,QHBoxLayout,QLabel,
    QPushButton,QLineEdit,QFileDialog,QTableWidget,QTableWidgetItem,QHeaderView,QAbstractItemView,
    QProgressBar,QCheckBox,QMessageBox,QFrame)
from bom_core import convert_one,validate_template

# PyQt's relocatable wheel may misdecode a Korean Windows installation path.
# Passing a Unicode plugin path explicitly avoids the corrupted auto-detected path.
QCoreApplication.setLibraryPaths([str(ROOT/'vendor/PyQt5/Qt5/plugins')])

class DropTable(QTableWidget):
    files_dropped=pyqtSignal(list)
    def __init__(self):
        super().__init__(0,5)
        self.setHorizontalHeaderLabels(['RAW 파일','제품코드','상태','BOM 행 수','처리 내용'])
        self.setAcceptDrops(True);self.viewport().setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DropOnly)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.verticalHeader().hide();self.setAlternatingRowColors(True)
        self.horizontalHeader().setSectionResizeMode(0,QHeaderView.Stretch)
        self.horizontalHeader().setSectionResizeMode(4,QHeaderView.Stretch)
        for col,width in [(1,100),(2,90),(3,95)]:self.setColumnWidth(col,width)
    def dragEnterEvent(self,event):
        if self.isEnabled() and event.mimeData().hasUrls():event.acceptProposedAction()
        else:event.ignore()
    def dragMoveEvent(self,event):self.dragEnterEvent(event)
    def dropEvent(self,event):
        if self.isEnabled() and event.mimeData().hasUrls():
            self.files_dropped.emit([url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()])
            event.acceptProposedAction()
        else:event.ignore()

class BatchWorker(QThread):
    started_file=pyqtSignal(int)
    result=pyqtSignal(int,dict)
    progress=pyqtSignal(int,int)
    completed=pyqtSignal(dict)
    def __init__(self,files,template,output,overwrite=False):
        super().__init__();self.files=list(files);self.template=template;self.output=Path(output);self.overwrite=overwrite
    def run(self):
        results=[];seen=set();totals={'완료':0,'실패':0,'건너뜀':0,'취소':0}
        for i,raw in enumerate(self.files):
            if self.isInterruptionRequested():
                result={'status':'취소','rows':0,'output':'','message':'사용자가 대기를 취소했습니다.'}
            elif Path(raw).stem in seen:
                result={'status':'건너뜀','rows':0,'output':'','message':'같은 제품코드가 목록에 중복되었습니다.'}
            else:
                self.started_file.emit(i)
                try:
                    result=convert_one(raw,self.template,self.output,self.overwrite)
                    if result['status'] in ['완료','건너뜀']:seen.add(Path(raw).stem)
                except Exception as exc:
                    result={'status':'실패','rows':0,'output':'','message':str(exc)}
            totals[result['status']]+=1
            results.append({'source':str(raw),**result});self.result.emit(i,result);self.progress.emit(i+1,len(self.files))
        log='';log_error=''
        try:
            log_dir=self.output/'변환기록';log_dir.mkdir(parents=True,exist_ok=True)
            log=log_dir/('BOM_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'.csv')
            with log.open('w',encoding='utf-8-sig',newline='') as stream:
                writer=csv.writer(stream);writer.writerow(['원본 파일','상태','BOM 행 수','완료 파일','처리 내용'])
                for r in results:writer.writerow([r['source'],r['status'],r['rows'],r['output'],r['message']])
        except Exception as exc:log_error=str(exc)
        self.completed.emit({'totals':totals,'log':str(log),'log_error':log_error})

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__();self.paths=[];self.worker=None;self.busy=False;self.last_summary=None
        self.setWindowTitle('BOM 일괄 변환기');self.resize(1120,800);self.setMinimumSize(900,650)
        settings={}
        try:settings=json.loads((ROOT/'settings.json').read_text(encoding='utf-8'))
        except (OSError,ValueError):pass
        self.template=QLineEdit(settings.get('template',str(ROOT/'templates/BOM 자동 입력 양식.xlsx')))
        self.output=QLineEdit(settings.get('output',str(ROOT/'완료파일')))
        central=QWidget();self.setCentralWidget(central);layout=QVBoxLayout(central);layout.setContentsMargins(28,24,28,22);layout.setSpacing(14)
        heading=QHBoxLayout();title=QLabel('BOM 일괄 변환');title.setObjectName('title');heading.addWidget(title);heading.addStretch()
        tag=QLabel('구 ERP → 신 ERP');tag.setObjectName('tag');heading.addWidget(tag);layout.addLayout(heading)
        subtitle=QLabel('RAW 파일을 추가하고 변환하면 제품코드별 완료 파일이 저장됩니다.');subtitle.setObjectName('muted');layout.addWidget(subtitle)
        self.template_button=QPushButton('양식 선택');self.template_button.clicked.connect(self.choose_template)
        self.output_button=QPushButton('폴더 선택');self.output_button.clicked.connect(self.choose_output)
        for label,field,button in [('빈 양식',self.template,self.template_button),('저장 폴더',self.output,self.output_button)]:
            line=QHBoxLayout();name=QLabel(label);name.setFixedWidth(74);line.addWidget(name);line.addWidget(field);line.addWidget(button);layout.addLayout(line)
        drop=QLabel('RAW 엑셀 파일 여러 개 또는 폴더를 아래 목록으로 끌어 놓으세요.');drop.setObjectName('dropHint');layout.addWidget(drop)
        tools=QHBoxLayout();self.add_button=QPushButton('파일 추가');self.folder_button=QPushButton('폴더 추가');self.remove_button=QPushButton('선택 삭제');self.clear_button=QPushButton('목록 비우기')
        self.add_button.clicked.connect(self.choose_files);self.folder_button.clicked.connect(self.choose_folder)
        self.remove_button.clicked.connect(self.remove_selected);self.clear_button.clicked.connect(self.clear_files)
        for button in [self.add_button,self.folder_button,self.remove_button,self.clear_button]:tools.addWidget(button)
        tools.addStretch();self.count_label=QLabel('대기 0개');tools.addWidget(self.count_label);layout.addLayout(tools)
        self.table=DropTable();self.table.files_dropped.connect(self.add_paths);layout.addWidget(self.table,1)
        options=QHBoxLayout();self.overwrite=QCheckBox('같은 이름의 완료 파일 덮어쓰기');options.addWidget(self.overwrite);options.addStretch()
        options.addWidget(QLabel('출력명: 제품코드 완료.xlsx'));layout.addLayout(options)
        rules=QLabel('G열 = RAW Z열 단위  ·  B열 = KG  ·  용도 = M  ·  대체번호 = 0  ·  BOM수량 = 1\nScrap은 기존 테스트 규칙인 직속 하위 구성품 개수로 입력됩니다.');rules.setObjectName('rules');layout.addWidget(rules)
        self.progress=QProgressBar();self.progress.setRange(0,1);self.progress.setValue(0);self.progress.setFormat('0 / 0');layout.addWidget(self.progress)
        actions=QHBoxLayout();self.status=QLabel('RAW 파일을 추가해 주세요.');self.status.setWordWrap(True);actions.addWidget(self.status,1)
        self.open_button=QPushButton('저장 폴더 열기');self.open_button.clicked.connect(self.open_output);actions.addWidget(self.open_button)
        self.stop_button=QPushButton('대기 취소');self.stop_button.setEnabled(False);self.stop_button.clicked.connect(self.stop_batch);actions.addWidget(self.stop_button)
        self.convert_button=QPushButton('변환 시작');self.convert_button.setObjectName('primary');self.convert_button.clicked.connect(self.start_batch);actions.addWidget(self.convert_button);layout.addLayout(actions)
        self.setStyleSheet('''
          QWidget {font-family:"Malgun Gothic";font-size:13px;color:#23344a;background:#f5f7fb;}
          QLabel#title {font-size:27px;font-weight:700;color:#142841;}
          QLabel#tag {background:#e3edfc;color:#275ca0;border-radius:10px;padding:7px 12px;}
          QLabel#muted {color:#67768c;}
          QLabel#dropHint {background:#eaf1fc;color:#345b91;border:1px dashed #9db7de;border-radius:8px;padding:14px;}
          QLabel#rules {color:#64748b;font-size:12px;}
          QLineEdit {background:white;border:1px solid #d2dce8;border-radius:6px;padding:9px;}
          QPushButton {background:white;border:1px solid #cfd9e5;border-radius:6px;padding:9px 15px;}
          QPushButton:hover {background:#eaf1fc;border-color:#87a8d4;}
          QPushButton:disabled {color:#9ba7b7;background:#edf0f5;}
          QPushButton#primary {background:#255fc0;color:white;border:none;font-weight:bold;padding:11px 25px;}
          QPushButton#primary:disabled {background:#94acce;}
          QTableWidget {background:white;alternate-background-color:#f7f9fc;border:1px solid #d6dfeb;border-radius:6px;gridline-color:#e8edf4;selection-background-color:#dbe9ff;selection-color:#193d74;}
          QHeaderView::section {background:#edf2f8;border:none;border-bottom:1px solid #d6dfeb;padding:10px;font-weight:bold;}
          QProgressBar {background:#e4ebf5;border:none;border-radius:5px;text-align:center;height:20px;}
          QProgressBar::chunk {background:#86afea;border-radius:5px;}
        ''')
        self.update_count()
    def choose_files(self):
        files,_=QFileDialog.getOpenFileNames(self,'RAW 파일 추가','','Excel RAW (*.xls *.xlsx)');self.add_paths(files)
    def choose_folder(self):
        folder=QFileDialog.getExistingDirectory(self,'RAW 폴더 추가')
        if folder:self.add_paths([folder])
    def choose_template(self):
        filename,_=QFileDialog.getOpenFileName(self,'빈 BOM 양식 선택',self.template.text(),'Excel (*.xlsx)')
        if filename:self.template.setText(filename)
    def choose_output(self):
        folder=QFileDialog.getExistingDirectory(self,'완료 파일 저장 폴더',self.output.text())
        if folder:self.output.setText(folder)
    def add_paths(self,paths):
        if self.busy:return
        known={os.path.normcase(str(p)) for p in self.paths};added=0;ignored=0
        self.table.setUpdatesEnabled(False)
        try:
            for name in paths:
                p=Path(name)
                candidates=sorted(p.rglob('*')) if p.is_dir() else [p]
                for candidate in candidates:
                    if not candidate.is_file() or candidate.suffix.lower() not in ['.xls','.xlsx']:continue
                    if not candidate.stem.isascii() or not candidate.stem.isdigit():ignored+=1;continue
                    resolved=candidate.resolve();key=os.path.normcase(str(resolved))
                    if key in known:continue
                    known.add(key);self.paths.append(resolved);row=self.table.rowCount();self.table.insertRow(row)
                    for col,text in enumerate([resolved.name,resolved.stem,'대기','—','']):
                        item=QTableWidgetItem(text);item.setToolTip(str(resolved));self.table.setItem(row,col,item)
                    added+=1
        finally:self.table.setUpdatesEnabled(True)
        self.update_count();self.status.setText(f'{added}개 추가됨'+(f' · 제품코드 파일명이 아닌 {ignored}개 제외' if ignored else ''))
    def update_count(self):
        self.count_label.setText(f'전체 {len(self.paths):,}개');self.convert_button.setEnabled(bool(self.paths) and not self.busy)
    def remove_selected(self):
        for index in sorted({x.row() for x in self.table.selectedIndexes()},reverse=True):self.table.removeRow(index);self.paths.pop(index)
        self.update_count()
    def clear_files(self):self.table.setRowCount(0);self.paths=[];self.update_count()
    def set_busy(self,busy):
        self.busy=busy
        for widget in [self.add_button,self.folder_button,self.remove_button,self.clear_button,self.template_button,self.output_button,self.template,self.output,self.overwrite]:widget.setEnabled(not busy)
        self.table.setAcceptDrops(not busy);self.table.viewport().setAcceptDrops(not busy)
        self.stop_button.setEnabled(busy);self.update_count()
    def start_batch(self):
        if self.busy or not self.paths:return
        try:
            if not self.template.text().strip() or not self.output.text().strip():raise ValueError('양식과 저장 폴더를 선택해 주세요.')
            template=Path(self.template.text()).resolve();output=Path(self.output.text()).resolve()
            validate_template(template);output.mkdir(parents=True,exist_ok=True)
            # A real write check catches protected output folders before starting a batch.
            import tempfile
            with tempfile.TemporaryFile(dir=output):pass
        except Exception as exc:QMessageBox.warning(self,'설정 확인',str(exc));return
        try:(ROOT/'settings.json').write_text(json.dumps({'template':str(template),'output':str(output)},ensure_ascii=False,indent=2),encoding='utf-8')
        except OSError:pass
        for row in range(self.table.rowCount()):
            for col,value in [(2,'대기'),(3,'—'),(4,'')]:self.table.item(row,col).setText(value)
        self.set_busy(True);self.progress.setRange(0,len(self.paths));self.progress.setValue(0)
        self.progress.setFormat(f'0 / {len(self.paths)}');self.status.setText('변환을 시작합니다.');self.last_summary=None
        self.worker=BatchWorker(self.paths,template,output,self.overwrite.isChecked())
        self.worker.started_file.connect(self.on_started);self.worker.result.connect(self.on_result)
        self.worker.progress.connect(self.on_progress);self.worker.completed.connect(self.on_completed)
        self.worker.finished.connect(self.on_finished);self.worker.start()
    def on_started(self,row):self.table.item(row,2).setText('변환 중');self.status.setText(f'{self.paths[row].name} 변환 중…')
    def on_result(self,row,result):
        for col,value in [(2,result['status']),(3,str(result['rows']) if result['rows'] else '—'),(4,result['message'])]:
            self.table.item(row,col).setText(value);self.table.item(row,col).setToolTip(result['message']+'\n'+result['output'])
    def on_progress(self,done,total):self.progress.setValue(done);self.progress.setFormat(f'{done:,} / {total:,}')
    def on_completed(self,summary):
        self.last_summary=summary;t=summary['totals']
        self.status.setText(f"완료 {t['완료']} · 실패 {t['실패']} · 건너뜀 {t['건너뜀']} · 취소 {t['취소']}"+(' · 기록 저장 실패: '+summary['log_error'] if summary['log_error'] else ''))
        self.status.setToolTip(summary['log'])
    def on_finished(self):self.set_busy(False)
    def stop_batch(self):
        if self.worker and self.worker.isRunning():self.worker.requestInterruption();self.stop_button.setEnabled(False);self.status.setText('현재 파일 저장 후 남은 대기를 취소합니다.')
    def open_output(self):
        folder=Path(self.output.text())
        if folder.is_dir():QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder.resolve())))
        else:QMessageBox.information(self,'저장 폴더','변환을 시작하면 저장 폴더가 생성됩니다.')
    def closeEvent(self,event):
        if self.busy:
            self.stop_batch();self.status.setText('현재 파일을 마무리하고 있습니다. 완료 후 창을 닫아 주세요.');event.ignore()
        else:event.accept()

def main():
    app=QApplication(sys.argv);app.setFont(QFont('Malgun Gothic',10));window=MainWindow()
    if '--smoke-test' in sys.argv:
        window.show();QTimer.singleShot(200,app.quit)
    else:window.show()
    return app.exec_()

if __name__=='__main__':
    try:sys.exit(main())
    except Exception:
        (ROOT/'실행오류.log').write_text(traceback.format_exc(),encoding='utf-8')
        raise
