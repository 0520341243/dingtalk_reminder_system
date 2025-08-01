#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
钉钉定时提醒系统后端服务 - 修复版本
支持Excel文件管理、定时任务调度和钉钉消息推送
"""

import os
import secrets
from flask import Flask, session
from flask import Flask, redirect, url_for
import hashlib  # 添加这行
import json
import logging
import schedule
import time
import requests
import pandas as pd
from datetime import datetime, timedelta
import pytz
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename
from threading import Thread
import sqlite3
from pathlib import Path 
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from werkzeug.utils import secure_filename
from threading import Thread
import sqlite3
from pathlib import Path

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('reminder_system.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# 设置密钥用于session
app.secret_key = os.getenv('SECRET_KEY', secrets.token_hex(32))

# 默认登录凭据（建议通过环境变量配置）
DEFAULT_USERNAME = os.getenv('ADMIN_USERNAME', 'admin')
DEFAULT_PASSWORD = os.getenv('ADMIN_PASSWORD', 'admin123')  # 建议修改默认密码

# 配置
class Config:
    UPLOAD_FOLDER = 'uploads'
    TEMP_FOLDER = 'temp_reminders'
    ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB
    DINGTALK_WEBHOOK = os.getenv('DINGTALK_WEBHOOK', '')
    DB_PATH = 'reminder_system.db'
    
    # 默认星期到子表映射（用户可通过前端修改）
    DEFAULT_WEEKDAY_SHEET_MAP = {
        0: '子表1',  # 周一
        1: '子表1',  # 周二
        2: '子表1',  # 周三
        3: '子表1',  # 周四
        4: '子表2',  # 周五
        5: '子表5',  # 周六
        6: '子表5'   # 周日
    }

app.config.from_object(Config)

# 确保目录存在
for folder in [Config.UPLOAD_FOLDER, Config.TEMP_FOLDER]:
    os.makedirs(folder, exist_ok=True)

class DatabaseManager:
    """数据库管理器"""
    
    def __init__(self, db_path):
        self.db_path = db_path
        self.init_db()
    
    def init_db(self):
        """初始化数据库"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 创建用户表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username VARCHAR(50) UNIQUE NOT NULL,
                password_hash VARCHAR(255) NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                last_login TIMESTAMP
            )
        ''')
        
        # 创建提醒任务表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS daily_reminders (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                time TEXT NOT NULL,
                message TEXT NOT NULL,
                sent BOOLEAN DEFAULT FALSE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 创建系统日志表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS system_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                level TEXT NOT NULL,
                message TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 创建配置表
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS system_config (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # 初始化默认用户
        self.init_default_user()
        
        conn.commit()
        conn.close()
    
    def init_default_user(self):
        """初始化默认用户"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 检查是否已有用户
        cursor.execute('SELECT COUNT(*) FROM users')
        user_count = cursor.fetchone()[0]
        
        if user_count == 0:
            # 创建默认管理员用户
            password_hash = self.hash_password(DEFAULT_PASSWORD)
            cursor.execute('''
                INSERT INTO users (username, password_hash)
                VALUES (?, ?)
            ''', (DEFAULT_USERNAME, password_hash))
            logger.info(f"已创建默认用户: {DEFAULT_USERNAME}")
        
        conn.commit()
        conn.close()
    
    def hash_password(self, password):
        """密码哈希"""
        return hashlib.sha256(password.encode()).hexdigest()
    
    def verify_user(self, username, password):
        """验证用户凭据"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        password_hash = self.hash_password(password)
        cursor.execute('''
            SELECT id FROM users 
            WHERE username = ? AND password_hash = ?
        ''', (username, password_hash))
        
        user = cursor.fetchone()
        
        if user:
            # 更新最后登录时间
            cursor.execute('''
                UPDATE users SET last_login = CURRENT_TIMESTAMP 
                WHERE id = ?
            ''', (user[0],))
            conn.commit()
        
        conn.close()
        return user is not None
    
    def change_password(self, username, old_password, new_password):
        """修改密码"""
        if not self.verify_user(username, old_password):
            return False
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        new_password_hash = self.hash_password(new_password)
        cursor.execute('''
            UPDATE users SET password_hash = ? 
            WHERE username = ?
        ''', (new_password_hash, username))
        
        conn.commit()
        conn.close()
        return True
    
    def save_daily_reminders(self, date, reminders, mark_past_as_sent=False):
        """保存当天提醒计划"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 删除当天已存在的提醒
        cursor.execute('DELETE FROM daily_reminders WHERE date = ?', (date,))
        
        # 获取当前时间（如果是今天的话）
        current_time = None
        if mark_past_as_sent and date == datetime.now().strftime('%Y-%m-%d'):
            current_time = datetime.now().strftime('%H:%M:%S')
        
        # 插入新的提醒
        for reminder in reminders:
            # 如果是今天且时间已过，标记为已发送
            sent_status = False
            if current_time and reminder['time'] <= current_time:
                sent_status = True
                logger.info(f"过去时间的提醒已标记为已发送: {reminder['time']} - {reminder['message'][:30]}...")
            
            cursor.execute('''
                INSERT INTO daily_reminders (date, time, message, sent)
                VALUES (?, ?, ?, ?)
            ''', (date, reminder['time'], reminder['message'], sent_status))
        
        conn.commit()
        conn.close()
        
        sent_count = len([r for r in reminders if current_time and r['time'] <= current_time])
        logger.info(f"保存了 {len(reminders)} 条提醒计划到数据库，日期：{date}，其中 {sent_count} 条标记为已发送")
    
    def get_daily_reminders(self, date):
        """获取当天提醒计划"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT id, time, message, sent FROM daily_reminders 
            WHERE date = ? ORDER BY time
        ''', (date,))
        
        reminders = []
        for row in cursor.fetchall():
            reminders.append({
                'id': row[0],
                'time': row[1],
                'message': row[2],
                'sent': bool(row[3])
            })
        
        conn.close()
        return reminders
    
    def mark_reminder_sent(self, reminder_id):
        """标记提醒已发送"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            UPDATE daily_reminders SET sent = TRUE 
            WHERE id = ?
        ''', (reminder_id,))
        
        conn.commit()
        conn.close()
    
    def log_message(self, level, message):
        """记录系统日志"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            INSERT INTO system_logs (level, message)
            VALUES (?, ?)
        ''', (level, message))
        
        conn.commit()
        conn.close()
    
    def get_config(self, key, default_value=None):
        """获取配置项"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('SELECT value FROM system_config WHERE key = ?', (key,))
        result = cursor.fetchone()
        
        conn.close()
        
        if result:
            try:
                # 尝试解析JSON格式的配置
                return json.loads(result[0])
            except:
                # 如果不是JSON，直接返回字符串值
                return result[0]
        
        return default_value
    
    def set_config(self, key, value):
        """设置配置项"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 如果值是字典或列表，转换为JSON字符串
        if isinstance(value, (dict, list)):
            value_str = json.dumps(value, ensure_ascii=False)
        else:
            value_str = str(value)
        
        cursor.execute('''
            INSERT OR REPLACE INTO system_config (key, value, updated_at)
            VALUES (?, ?, CURRENT_TIMESTAMP)
        ''', (key, value_str))
        
        conn.commit()
        conn.close()
        
        logger.info(f"配置已更新: {key}")
    
    def get_weekday_sheet_map(self):
        """获取星期-工作表映射配置"""
        return self.get_config('weekday_sheet_map', Config.DEFAULT_WEEKDAY_SHEET_MAP)

class ExcelProcessor:
    """Excel文件处理器"""
    
    @staticmethod
    def allowed_file(filename):
        """检查文件是否允许上传"""
        return '.' in filename and \
               filename.rsplit('.', 1)[1].lower() in Config.ALLOWED_EXTENSIONS
    
    @staticmethod
    def read_excel_reminders(file_path, sheet_name):
        """读取Excel文件中的提醒数据"""
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # 检查必要的列是否存在
            if '时间' not in df.columns or '消息内容' not in df.columns:
                raise ValueError(f"Excel文件中的'{sheet_name}'表缺少必要的列：时间、消息内容")
            
            reminders = []
            for _, row in df.iterrows():
                time_str = str(row['时间']).strip()
                message = str(row['消息内容']).strip()
                
                if pd.isna(row['时间']) or pd.isna(row['消息内容']):
                    continue
                
                # 处理时间格式
                if isinstance(row['时间'], pd.Timestamp):
                    time_str = row['时间'].strftime('%H:%M:%S')
                elif ':' not in time_str:
                    continue
                
                reminders.append({
                    'time': time_str,
                    'message': message
                })
            
            return reminders
        except Exception as e:
            logger.error(f"读取Excel文件失败: {e}")
            raise

class DingTalkNotifier:
    """钉钉消息发送器"""
    
    def __init__(self, webhook_url):
        self.webhook_url = webhook_url
    
    def send_message(self, message):
        """发送钉钉消息"""
        if not self.webhook_url:
            logger.error("钉钉Webhook地址未配置")
            return False
        
        payload = {
            "msgtype": "text",
            "text": {
                "content": message
            }
        }
        
        try:
            response = requests.post(
                self.webhook_url,
                json=payload,
                timeout=10
            )
            
            if response.status_code == 200:
                result = response.json()
                if result.get('errcode') == 0:
                    logger.info(f"消息发送成功: {message[:50]}...")
                    return True
                else:
                    logger.error(f"钉钉返回错误: {result}")
                    return False
            else:
                logger.error(f"HTTP请求失败: {response.status_code}")
                return False
                
        except Exception as e:
            logger.error(f"发送钉钉消息异常: {e}")
            return False

class ReminderScheduler:
    """提醒任务调度器"""
    
    def __init__(self):
        self.db_manager = DatabaseManager(Config.DB_PATH)
        self.notifier = DingTalkNotifier(Config.DINGTALK_WEBHOOK)
        self.is_running = False
        self.scheduler_thread = None
    
    def load_daily_plan(self, mark_past_as_sent=False, is_immediate=False):
        """加载当天提醒计划
        
        Args:
            mark_past_as_sent: 是否将过去时间的提醒标记为已发送
            is_immediate: 是否立即生效（用于临时计划）
        """
        today = datetime.now().strftime('%Y-%m-%d')
        weekday = datetime.now().weekday()
        
        try:
            # 检查是否有临时提醒文件
            temp_files = list(Path(Config.TEMP_FOLDER).glob('*.xlsx'))
            temp_files.extend(list(Path(Config.TEMP_FOLDER).glob('*.xls')))
            
            if temp_files:
                # 使用最新的临时文件
                temp_file = max(temp_files, key=os.path.getctime)
                logger.info(f"发现临时提醒文件: {temp_file}")
                
                # 读取临时文件的第一个工作表
                reminders = ExcelProcessor.read_excel_reminders(temp_file, 0)
                self.db_manager.save_daily_reminders(today, reminders, mark_past_as_sent)
                
                # 如果不是立即生效，则备份并删除临时文件
                if not is_immediate:
                    backup_name = f"temp_backup_{today}_{temp_file.name}"
                    temp_file.rename(Path(Config.UPLOAD_FOLDER) / backup_name)
                    logger.info(f"临时文件已备份为: {backup_name}")
                
            else:
                # 使用常规Excel文件
                excel_files = list(Path(Config.UPLOAD_FOLDER).glob('*.xlsx'))
                excel_files.extend(list(Path(Config.UPLOAD_FOLDER).glob('*.xls')))
                
                if not excel_files:
                    logger.warning("未找到Excel提醒文件")
                    return 0
                
                # 使用最新的Excel文件
                excel_file = max(excel_files, key=os.path.getctime)
                
                # 获取用户自定义的星期-工作表映射
                weekday_map = self.db_manager.get_weekday_sheet_map()
                sheet_name = weekday_map.get(str(weekday), weekday_map.get(weekday, '子表1'))
                
                logger.info(f"使用Excel文件: {excel_file}, 工作表: {sheet_name}, 星期: {weekday}")
                reminders = ExcelProcessor.read_excel_reminders(excel_file, sheet_name)
                self.db_manager.save_daily_reminders(today, reminders, mark_past_as_sent)
            
            # 记录加载结果
            status = "立即生效" if is_immediate else "定时生效"
            past_handled = "已跳过过去时间" if mark_past_as_sent else "包含所有时间"
            self.db_manager.log_message('INFO', f'成功加载 {len(reminders)} 条提醒计划 ({status}, {past_handled})')
            
            return len(reminders)
            
        except Exception as e:
            error_msg = f"加载提醒计划失败: {e}"
            logger.error(error_msg)
            self.db_manager.log_message('ERROR', error_msg)
            return 0
    
    def check_and_send_reminders(self):
        """检查并发送到时的提醒"""
        today = datetime.now().strftime('%Y-%m-%d')
        current_time = datetime.now().strftime('%H:%M:%S')
        
        reminders = self.db_manager.get_daily_reminders(today)
        
        for reminder in reminders:
            if not reminder['sent'] and reminder['time'] <= current_time:
                success = self.notifier.send_message(reminder['message'])
                if success:
                    self.db_manager.mark_reminder_sent(reminder['id'])
                    self.db_manager.log_message('INFO', f"提醒已发送: {reminder['message'][:50]}...")
                else:
                    self.db_manager.log_message('ERROR', f"提醒发送失败: {reminder['message'][:50]}...")
    
    def start_scheduler(self):
        """启动调度器"""
        if self.is_running:
            return
        
        self.is_running = True
        
        # 设置定时任务
        schedule.every().day.at("02:00").do(lambda: self.load_daily_plan())
        schedule.every().minute.do(self.check_and_send_reminders)
        
        # 立即加载当天计划（启动时跳过过去时间的提醒）
        self.load_daily_plan(mark_past_as_sent=True)
        
        def run_scheduler():
            while self.is_running:
                schedule.run_pending()
                time.sleep(1)
        
        self.scheduler_thread = Thread(target=run_scheduler, daemon=True)
        self.scheduler_thread.start()
        
        logger.info("提醒调度器已启动")
    
    def stop_scheduler(self):
        """停止调度器"""
        self.is_running = False
        schedule.clear()
        logger.info("提醒调度器已停止")

# 全局调度器实例
scheduler = ReminderScheduler()

# 登录检查装饰器
def login_required(f):
    def decorated_function(*args, **kwargs):
        if 'logged_in' not in session:
            if request.is_json:
                return jsonify({'error': '请先登录', 'login_required': True}), 401
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    decorated_function.__name__ = f.__name__
    return decorated_function

# Flask路由定义
@app.route('/login', methods=['GET', 'POST'])
def login():
    """登录页面"""
    if request.method == 'POST':
        if request.is_json:
            data = request.get_json()
            username = data.get('username', '')
            password = data.get('password', '')
        else:
            username = request.form.get('username', '')
            password = request.form.get('password', '')
        
        if scheduler.db_manager.verify_user(username, password):
            session['logged_in'] = True
            session['username'] = username
            logger.info(f"用户登录成功: {username}")
            
            if request.is_json:
                return jsonify({'success': True, 'message': '登录成功'})
            return redirect(url_for('index'))
        else:
            logger.warning(f"登录失败: {username}")
            error_msg = '用户名或密码错误'
            
            if request.is_json:
                return jsonify({'error': error_msg}), 401
            return render_template('login.html', error=error_msg)
    
    # 如果已登录，直接跳转到主页
    if 'logged_in' in session:
        return redirect(url_for('index'))
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    """退出登录"""
    username = session.get('username', 'Unknown')
    session.clear()
    logger.info(f"用户退出登录: {username}")
    return redirect(url_for('login'))

@app.route('/')
@login_required
def index():
    """首页"""
    return render_template('index.html', username=session.get('username'))

@app.route('/change_password', methods=['POST'])
@login_required 
def change_password():
    """修改密码"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': '请求体不能为空或不是JSON格式'}), 400

        old_password = data.get('old_password', '')
        new_password = data.get('new_password', '')

        if not old_password or not new_password:
            return jsonify({'error': '请填写完整的密码信息'}), 400

        if len(new_password) < 6:
            return jsonify({'error': '新密码长度至少6位'}), 400

        username = session.get('username')
        if not username:
            return jsonify({'error': '用户未登录或会话已过期'}), 401

        if scheduler.db_manager.change_password(username, old_password, new_password):
            logger.info(f"用户修改密码成功: {username}")
            return jsonify({'message': '密码修改成功'})
        else:
            return jsonify({'error': '原密码错误'}), 400

    except Exception as e:
        logger.error(f"修改密码失败: {e}")
        return jsonify({'error': '服务器内部错误，请稍后再试'}), 500

@app.route('/api/upload', methods=['POST'])
@login_required
def upload_file():
    """上传Excel文件"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': '没有选择文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400
        
        if file and ExcelProcessor.allowed_file(file.filename):
            filename = secure_filename(file.filename)
            
            # 添加时间戳避免文件名冲突
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            name, ext = os.path.splitext(filename)
            filename = f"{name}_{timestamp}{ext}"
            
            file_path = os.path.join(Config.UPLOAD_FOLDER, filename)
            file.save(file_path)
            
            logger.info(f"文件上传成功: {filename}")
            return jsonify({'message': '文件上传成功', 'filename': filename})
        
        return jsonify({'error': '不支持的文件格式'}), 400
    except Exception as e:
        logger.error(f"文件上传异常: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/upload_temp', methods=['POST'])
@login_required
def upload_temp_file():
    """上传临时提醒文件"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': '没有选择文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400
        
        # 获取是否立即启动参数
        immediate_start = request.form.get('immediate', 'false').lower() == 'true'
        
        if file and ExcelProcessor.allowed_file(file.filename):
            # 清理现有临时文件
            for temp_file in Path(Config.TEMP_FOLDER).glob('*'):
                temp_file.unlink()
            
            filename = secure_filename(file.filename)
            file_path = os.path.join(Config.TEMP_FOLDER, filename)
            file.save(file_path)
            
            message = f"临时文件上传成功: {filename}"
            
            if immediate_start:
                try:
                    # 立即加载临时计划
                    count = scheduler.load_daily_plan(mark_past_as_sent=True, is_immediate=True)
                    message += f"，已立即生效并加载 {count} 条提醒"
                    logger.info(f"临时计划立即启动: {filename}, 加载 {count} 条提醒")
                except Exception as e:
                    logger.error(f"临时计划立即启动失败: {e}")
                    message += f"，但立即启动失败: {str(e)}"
            else:
                message += "，将在明天2点生效"
            
            logger.info(f"临时文件上传: {filename}, 立即启动: {immediate_start}")
            return jsonify({'message': message, 'filename': filename, 'immediate': immediate_start})
        
        return jsonify({'error': '不支持的文件格式'}), 400
    except Exception as e:
        logger.error(f"临时文件上传异常: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/reminders/today')
def get_today_reminders():
    """获取今天的提醒计划"""
    try:
        today = datetime.now().strftime('%Y-%m-%d')
        reminders = scheduler.db_manager.get_daily_reminders(today)
        return jsonify({
            'date': today,
            'reminders': reminders,
            'total': len(reminders),
            'sent': len([r for r in reminders if r['sent']])
        })
    except Exception as e:
        logger.error(f"获取今日提醒失败: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/scheduler/start', methods=['POST'])
def start_scheduler():
    """启动调度器"""
    try:
        scheduler.start_scheduler()
        return jsonify({'message': '调度器已启动'})
    except Exception as e:
        logger.error(f"启动调度器失败: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/scheduler/stop', methods=['POST'])
def stop_scheduler():
    """停止调度器"""
    try:
        scheduler.stop_scheduler()
        return jsonify({'message': '调度器已停止'})
    except Exception as e:
        logger.error(f"停止调度器失败: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/scheduler/reload', methods=['POST'])
def reload_plan():
    """重新加载提醒计划"""
    try:
        count = scheduler.load_daily_plan(mark_past_as_sent=True)
        return jsonify({
            'message': f'提醒计划已重新加载，共 {count} 条提醒，已跳过过去时间的提醒',
            'count': count
        })
    except Exception as e:
        logger.error(f"重新加载计划失败: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/test_message', methods=['POST'])
def test_message():
    """测试钉钉消息发送"""
    try:
        data = request.get_json()
        message = data.get('message', '这是一条测试消息')
        
        success = scheduler.notifier.send_message(message)
        if success:
            return jsonify({'message': '测试消息发送成功'})
        else:
            return jsonify({'error': '测试消息发送失败'}), 500
    except Exception as e:
        logger.error(f"测试消息发送异常: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/temp_files/clear', methods=['POST'])
def clear_temp_files():
    """清理临时文件"""
    try:
        count = 0
        for temp_file in Path(Config.TEMP_FOLDER).glob('*'):
            temp_file.unlink()
            count += 1
        
        logger.info(f"清理了 {count} 个临时文件")
        return jsonify({'message': f'已清理 {count} 个临时文件'})
    except Exception as e:
        logger.error(f"清理临时文件失败: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/temp_files/status')
def temp_files_status():
    """获取临时文件状态"""
    try:
        temp_files = list(Path(Config.TEMP_FOLDER).glob('*.xlsx'))
        temp_files.extend(list(Path(Config.TEMP_FOLDER).glob('*.xls')))
        
        files_info = []
        for temp_file in temp_files:
            files_info.append({
                'name': temp_file.name,
                'size': temp_file.stat().st_size,
                'created': datetime.fromtimestamp(temp_file.stat().st_ctime).strftime('%Y-%m-%d %H:%M:%S')
            })
        
        return jsonify({
            'count': len(files_info),
            'files': files_info
        })
    except Exception as e:
        logger.error(f"获取临时文件状态失败: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/config', methods=['GET', 'POST'])
def manage_config():
    """管理系统配置"""
    if request.method == 'GET':
        try:
            weekday_map = scheduler.db_manager.get_weekday_sheet_map()
            return jsonify({
                'webhook_configured': bool(Config.DINGTALK_WEBHOOK),
                'scheduler_running': scheduler.is_running,
                'weekday_map': weekday_map
            })
        except Exception as e:
            logger.error(f"获取配置失败: {e}")
            return jsonify({'error': str(e)}), 500
    
    elif request.method == 'POST':
        try:
            data = request.get_json()
            if not data:
                return jsonify({'error': '请求数据为空'}), 400
            
            if 'dingtalk_webhook' in data:
                Config.DINGTALK_WEBHOOK = data['dingtalk_webhook']
                scheduler.notifier.webhook_url = data['dingtalk_webhook']
                scheduler.db_manager.set_config('dingtalk_webhook', data['dingtalk_webhook'])
            
            if 'weekday_map' in data:
                # 保存用户自定义的星期-工作表映射
                scheduler.db_manager.set_config('weekday_sheet_map', data['weekday_map'])
                logger.info(f"星期工作表映射已更新: {data['weekday_map']}")
            
            return jsonify({'message': '配置更新成功'})
        except Exception as e:
            logger.error(f"保存配置失败: {e}")
            return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    # 启动调度器
    scheduler.start_scheduler()
    
    # 启动Web服务
    app.run(host='0.0.0.0', port=5000, debug=False)