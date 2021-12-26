using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace SoulMouse
{
    public partial class Form1 : Form
    {
        // ホットキー関連の定数
        // モディファイアキー
        const int MOD_ALT = 0x0001;
        const int MOD_CONTROL = 0x0002;
        const int MOD_SHIFT = 0x0004;
        const int MOD_WIN = 0x0008;

        // ホットキーのイベントを示すメッセージID
        const int WM_HOTKEY = 0x0312;

        // ホットキー登録の際に指定するID
        // 0x0000〜0xbfff 内の適当な値を指定
        const int APP_START_HOTKEY_ID = 0x0001;
        const int APP_STOP_HOTKEY_ID = 0x0002;

        // ホットキーの登録
        [DllImport("user32.dll")]
        extern static int RegisterHotKey(IntPtr HWnd, int ID, int MOD_KEY, int KEY);

        // ホットキーの解除
        [DllImport("user32.dll")]
        extern static int UnregisterHotKey(IntPtr HWnd, int ID);


        public Form1()
        {
            InitializeComponent();

            // ホットキーの登録
            // CTRL + SHIFT + S　で機能開始
            if (RegisterHotKey(this.Handle, APP_START_HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, (int)Keys.S) == 0)
            {
                MessageBox.Show("「CTRL + SHIFT + S」をホットキーに登録できませんでした。");
                appExit();
            }
            // CTRL + SHIFT + D　で機能停止
            if (RegisterHotKey(this.Handle, APP_STOP_HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, (int)Keys.D) == 0)
            {
                MessageBox.Show("「CTRL + SHIFT + D」をホットキーに登録できませんでした。");
                appExit();
            }
        }

        // OSから渡されるメッセージ処理をオーバーライド
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);

            if (m.Msg == WM_HOTKEY)
            {
                if ((int)m.WParam == APP_START_HOTKEY_ID)
                {
                    // 機能開始
                    autoMouseRunStart();
                }
                else if ((int)m.WParam == APP_STOP_HOTKEY_ID)
                {
                    // 機能停止
                    autoMouseRunStop();
                }
            }
        }


        // マウスイベント関連の定数
        const int MOUSEEVENTF_MOVED = 0x0001;
        const int MOUSEEVENTF_ABSOLUTE = 0x8000;
        const int MOUSEEVENTF_WHEEL = 0x800;
        const int screen_length = 0x10000;

        // マウス操作を制御するためのMOUSEINPUT構造体
        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        }
        // SendInputメソッド用の構造体
        [StructLayout(LayoutKind.Sequential)]
        struct INPUT
        {
            public int type;
            public MOUSEINPUT mi;
        }
        // SendInputメソッドを宣言
        [DllImport("user32.dll")]
        extern static uint SendInput(
            uint nInputs,     // INPUT 構造体の数(イベント数)
            INPUT[] pInputs,  // INPUT 構造体
            int cbSize        // INPUT 構造体のサイズ
        );


        // 機能の開始／停止を制御するフラグ
        Boolean isAppRun = false;

        // 機能開始
        private async void autoMouseRunStart()
        {
            isAppRun = true;
            int screen_width = Screen.PrimaryScreen.Bounds.Width;
            int screen_height = Screen.PrimaryScreen.Bounds.Height;
            mStopwatch.Restart();
            mInput[0].mi.dx = Cursor.Position.X * (65535 / screen_width);
            mInput[0].mi.dy = Cursor.Position.Y * (65535 / screen_height);
            mInput[0].mi.dwFlags = MOUSEEVENTF_MOVED | MOUSEEVENTF_ABSOLUTE;

            // 開始時点でマウスがPrimaryScreenに無い場合、PrimaryScreenに強制移動させる
            if (Cursor.Position.X > screen_width || Cursor.Position.Y > screen_height)
            {
                mInput[0].mi.dx = (screen_width / 2) * (65535 / screen_width);
                mInput[0].mi.dy = (screen_height / 2) * (65535 / screen_height);
                SendInput(1, mInput, Marshal.SizeOf(mInput[0]));
            }

            await Task.Run(() =>
            {
                while (isAppRun)
                {
                    switch (mState)
                    {
                        case EState.Wait:
                            // 一定時間経過したら EState.Move に移行
                            if (mStopwatch.ElapsedMilliseconds > mWaitTime)
                            {
                                mState = EState.Move;
                                mMoveTime = mRandom.Next(1000, 3000); // 1～3秒でマウス移動
                                mTargetX = mRandom.Next(0, screen_width);  // ランダムなX座標
                                mTargetY = mRandom.Next(0, screen_height); // ランダムなY座標
                                mInput[0].mi.dwFlags = MOUSEEVENTF_MOVED | MOUSEEVENTF_ABSOLUTE;
                                mStopwatch.Restart();
                            }

                            // 適当にスリープさせる
                            // さもないと、CPU100%で処理してしまう
                            System.Threading.Thread.Sleep(5);
                            break;

                        case EState.Move:
                            // 一定時間経過したら EState.Wheel に移行
                            if (mStopwatch.ElapsedMilliseconds > mMoveTime)
                            {
                                mState = EState.Wheel;
                                mWheelTime = mRandom.Next(100, 800); // 0.1～0.8秒でマウスホイール
                                mInput[0].mi.dwFlags = MOUSEEVENTF_WHEEL;
                                mInput[0].mi.mouseData = mRandom.Next(-1, 2);
                                mStopwatch.Restart();
                            }

                            // カーソルを目標に追従させる
                            mInput[0].mi.dx += (int)((mTargetX - Cursor.Position.X) * 0.07) * (65535 / screen_width);
                            mInput[0].mi.dy += (int)((mTargetY - Cursor.Position.Y) * 0.07) * (65535 / screen_height);

                            // 適当にスリープさせる
                            // さもないと、カーソル移動処理が実行されすぎて移動が速くなりすぎてしまう
                            System.Threading.Thread.Sleep(5);

                            break;

                        case EState.Wheel:
                            // 一定時間経過したら EState.Wait に移行
                            if (mStopwatch.ElapsedMilliseconds > mWheelTime)
                            {
                                mState = EState.Wait;
                                mWaitTime = mRandom.Next(10000, 30000); // 10～300秒でマウス停止
                                mInput[0].mi.dwFlags = 0;
                                mStopwatch.Restart();
                            }

                            // マウスホイールではスリープさせない
                            // スリープさせると滑らかなスクロールにならなかった
                            //System.Threading.Thread.Sleep(1);

                            break;
                    }

                    // イベントの実行
                    // Cursor.Position への代入や、SetCursorPos での座標セットではスリープ状態になってしまうので SendInput を使用
                    SendInput(1, mInput, Marshal.SizeOf(mInput[0]));
                }
            });
        }

        // 機能停止
        private void autoMouseRunStop()
        {
            isAppRun = false;
            mState = EState.Wait;
            mTargetX = mTargetY = 0;
            mWaitTime = mMoveTime = mWheelTime = 2000;
        }

        // 終了メニュー押下
        private void AppExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            appExit();
        }

        // アプリケーションの終了
        private void appExit()
        {
            // ホットキーの解除
            UnregisterHotKey(this.Handle, APP_START_HOTKEY_ID);
            UnregisterHotKey(this.Handle, APP_STOP_HOTKEY_ID);

            // アプリ終了
            notifyIcon1.Visible = false;
            Application.Exit();
        }


        INPUT[] mInput = new INPUT[1];
        Stopwatch mStopwatch = new Stopwatch();
        enum EState { Wait, Move, Wheel }
        EState mState = EState.Wait;
        Random mRandom = new Random();
        int mTargetX = 0, mTargetY = 0;

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        int mWaitTime = 2000, mMoveTime = 2000, mWheelTime = 2000;
    }
}
