# connectors/hive_connection.py

import json
from pyhive import hive
from contextlib import contextmanager


def load_secrets():
    """
    从 secrets.json 文件加载敏感信息

    :return: 配置字典，如果加载失败则返回 None
    """
    try:
        with open('../secrets/secrets.json', 'r') as f:
            secrets = json.load(f)
        return secrets
    except Exception as e:
        print(f"读取 secrets.json 时出错: {e}")
        return None


@contextmanager
def get_hive_connection():
    """
    Hive 连接的上下文管理器，自动管理连接的生命周期

    :return: Hive 连接对象
    """
    secrets = load_secrets()

    if secrets is None:
        raise ValueError("无法加载 Hive 连接信息。请检查 secrets.json 文件。")

    hive_config = secrets.get('hive', {})
    if not hive_config:
        raise ValueError("无法从 secrets.json 获取 Hive 连接信息。")

    host = hive_config.get('host')
    port = hive_config.get('port', 10000)
    username = hive_config.get('username')
    password = hive_config.get('password')

    if not all([host, username, password]):
        raise ValueError("Hive 连接信息不完整，请确保 'host', 'username' 和 'password' 都已配置。")

    try:
        print(f"连接到 Hive: {host}:{port}...")
        connection = hive.Connection(
            host=host,
            port=port,
            username=username,
            password=password,
            auth='CUSTOM'
        )
        yield connection
    except Exception as e:
        print(f"连接到 Hive 时发生错误: {e}")
        raise
    finally:
        if connection:
            connection.close()
            print("Hive 连接已关闭。")

# 测试
if __name__ == "__main__":
    with get_hive_connection() as connection:
        cursor = connection.cursor()
        cursor.execute("SHOW DATABASES")
        print(cursor.fetchall())