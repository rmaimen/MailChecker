# -*- coding: utf-8 -*-

"""
    MailChecker.py

    IMAP4でメールサーバーに接続し、新着メールがない確認する。

    動作確認環境
        python:     3.8.6
        imapclient:     2.2.0
"""
import sys
import imaplib
import time
import _thread
import json
from imapclient import IMAPClient, exceptions
# Windows依存
import msvcrt
import winsound

# データの保存先
SETTING_FILE = "settings.json"

setting = {
    "server" : "",
    "port" : "",
    "ssl" : False,
    "user" : "",
    "password" : "",
    "max_retry" : 5,
	"wait_time" : 60
}

def main():
    """ エントリポイント

    """
    # 設定ファイルを読み出し
    with open(SETTING_FILE, "r") as f:
        temp = json.load(f)
        for key in temp:
            setting[key] = temp[key]

    args = sys.argv
    if len(args) == 2 and args[1] == 'push':
        push()
    else:
        polling()


# Pushで取得
def push():
    def connection():
        imap = IMAPClient(host=setting["server"], ssl=setting["ssl"], timeout=5)
        imap.login(setting["user"], setting["password"])
        imap.select_folder("INBOX")
        imap.idle()
        print("ideling...")
        return imap

    retry_count = 0
    is_input_list = []

    # 別スレッドでキー入力を待つ  
    _thread.start_new_thread(input_thread, (is_input_list,))
    print("プッシュ通知(push)で取得")
    print("終了するには、適当な文字を1文字入力してください。(例:Enter キーの押下)")
    try:
        imap = connection()

        while not is_input_list:
            try:
                print("\rcheck", end="")
                responses = imap.idle_check(timeout=setting["wait_time"])
                if len(responses) != 0:
                    for response in responses:
                        print("\rServer sent:", response)
                        if response[1] == b"EXISTS":
                            # 新着メールあり
                            print("\r新着メールがありました。")
                            winsound.PlaySound(r"C:\\Windows\\Media\\Windows Notify Email.wav", winsound.SND_FILENAME)
                        if response[0] == b"BYE":
                            # ログアウトされた
                            imap = connection()
            except (imaplib.IMAP4.abort, exceptions.ProtocolError) as e:
                print("接続が切れました。！".format(retry_count))
                print(e)
                imap = connection()
                retry_count += 1
                if retry_count <= setting["max_retry"]:
                    continue
                else:
                    print("リトライカウントオーバー:{0}".format(retry_count))
                    break
            # except KeyboardInterrupt:
            #     break
        imap.idle_done()
        print("IDLE mode done")
    finally:
        imap.logout()


# ポーリングで取得
def polling():
    def connection():
        imap = IMAPClient(host=setting["server"], ssl=setting["ssl"], timeout=5)
        imap.login(setting["user"], setting["password"])
        return imap

    retry_count = 0
    is_input_list = []

    # 別スレッドでキー入力を待つ  
    _thread.start_new_thread(input_thread, (is_input_list,))
    print("ポーリング(polling)で取得")
    print("終了するには、適当な文字を1文字入力してください。(例:Enter キーの押下)")
    try:
        imap = connection()

        # ポーリングで取得
        is_new_mail, new_count = check_new_mail(imap, 0, True)

        # キー入力がされていない間は処理を繰り返す
        i = 1
        while not is_input_list:
            try:
                print("\rチェック{0}回目({1}秒毎)".format(str(i), setting["wait_time"]), end="")

                is_new_mail, new_count = check_new_mail(imap, new_count, False)

                # 受信があった場合、音を鳴らす
                if is_new_mail:
                    print("\r新着メールがありました。")
                    winsound.PlaySound(r"C:\\Windows\\Media\\Windows Notify Email.wav", winsound.SND_FILENAME)
                
                i += 1
                # 次回確認までの待ち時間（サーバー負荷軽減のため）
                time.sleep(setting["wait_time"])

            except imaplib.IMAP4.abort as e:
                print("\r取得失敗(abort)！リトライ:{0}    Error:{1}".format(retry_count, e))
                imap = connection()
                retry_count += 1

                # リトライ上限を超えた場合、終了
                if retry_count > setting["max_retry"]:
                    break
                else:
                    continue
                
    except imaplib.IMAP4.error as e:
        print("\r取得失敗(error)！    Error:{0}".format(e))
    finally:
        imap.logout()


# キー入力の有無を記録する
def input_thread(is_input_list):
    msvcrt.getch()
    print("\rプログラムを終了中・・・")
    is_input_list.append(True)


# 新着メールが増えているか確認
def check_new_mail(imap, pre_new_count, is_first):
    # 新着メール数の取得
    new_count = 0
    imap.select_folder("INBOX", readonly=True)

    # 総数が増えたら処理を行う
    uids = imap.search("ALL")
    if len(uids) != 0:
        new_count = len(uids)

    is_new_mail = False
    if not is_first:
        # 初回実行時でない場合
        if new_count > pre_new_count:
            # 前回実行時より、新着カウントが増えている場合
            is_new_mail = True

    # # Debug
    # debug_str = "\tnew_count={0}    pre_new_count={1}   is_first={2}  >?={3}  is_new_mail={4}"
    # print(debug_str.format(new_count,pre_new_count,is_first,(new_count > pre_new_count),is_new_mail))

    return is_new_mail, new_count


if __name__ == "__main__":
    main()