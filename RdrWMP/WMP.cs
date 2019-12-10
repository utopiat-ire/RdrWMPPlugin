using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WMPLib;

namespace Produire.WMP
{
	[種類(DocUrl = "https://rdr.utopiat.net/docs/plugins/RdrWMP/wmp.htm"), メインスレッド]
	public class WindowsMediaPlayer : IProduireClass, IDisposable
	{
		WMPLib.WindowsMediaPlayer mediaPlayer = new WMPLib.WindowsMediaPlayer();
		メディア情報 media;

		#region イベント手順

		public event EventHandler 状態が変更された;
		public event EventHandler 再生状態が変更された;
		public event EventHandler 現在メディアが変更された;
		public event EventHandler メディア情報が変更された;

		#endregion

		public WindowsMediaPlayer()
		{
			SetEventHandling();
			mediaPlayer.settings.autoStart = false;
		}

		#region 手順

		[自分("を")]
		public void 再生()
		{
			try
			{
				mediaPlayer.controls.play();
			}
			catch (COMException ex)
			{
				throw new ProduireException("再生できない状態です。HRESULT:" + ex.ErrorCode.ToString("X"));
			}
		}
		[自分("を")]
		public void 一時停止()
		{
			try
			{
				mediaPlayer.controls.pause();
			}
			catch (COMException ex)
			{
				throw new ProduireException("一時停止が利用できない状態です。HRESULT:" + ex.ErrorCode.ToString("X"));
			}
		}
		[自分("を")]
		public void 停止()
		{
			try
			{
				mediaPlayer.controls.stop();
			}
			catch (COMException ex)
			{
				throw new ProduireException("停止が利用できない状態です。HRESULT:" + ex.ErrorCode.ToString("X"));
			}
		}
		[自分("を"), 補語("次へ"), 動詞("進む")]
		public void 次へ進む()
		{
			try
			{
				mediaPlayer.controls.next();
			}
			catch (COMException ex)
			{
				throw new ProduireException("この操作が利用できない状態です。HRESULT:" + ex.ErrorCode.ToString("X"));
			}
		}
		[自分("を"), 補語("前へ"), 動詞("戻る")]
		public void 前へ戻る()
		{
			try
			{ mediaPlayer.controls.previous(); }
			catch (COMException ex)
			{
				throw new ProduireException("この操作が利用できない状態です。HRESULT:" + ex.ErrorCode.ToString("X"));
			}
		}

		#endregion

		#region プロパティ

		[非表示]
		public string アドレス
		{
			get { return mediaPlayer.URL; }
			set
			{
				try
				{
					mediaPlayer.URL = value;
				}
				catch (COMException ex)
				{
					throw new ProduireException("ファイルを開くことができません。HRESULT:" + ex.ErrorCode.ToString("X"));
				}
			}
		}
		[非表示]
		public double 再生位置
		{
			get { return mediaPlayer.controls.currentPosition; }
			set
			{
				try
				{
					mediaPlayer.controls.currentPosition = value;
				}
				catch (COMException ex)
				{
					throw new ProduireException("現在の再生状態では設定できません。HRESULT:" + ex.ErrorCode.ToString("X"));
				}
			}
		}
		[非表示]
		public string 再生位置情報
		{
			get { return mediaPlayer.controls.currentPositionString; }
		}
		[非表示]
		public メディア情報 現在メディア
		{
			get { return media; }
			set
			{
				try
				{
					mediaPlayer.currentMedia = value.media;
				}
				catch (COMException ex)
				{
					throw new ProduireException("現在の再生状態では設定できません。HRESULT:" + ex.ErrorCode.ToString("X"));
				}
			}
		}
		public bool 自動再生
		{
			get { return mediaPlayer.settings.autoStart; }
			set { mediaPlayer.settings.autoStart = value; }
		}
		public int バランス
		{
			get { return mediaPlayer.settings.balance; }
			set { mediaPlayer.settings.balance = value; }
		}
		public bool 消音
		{
			get { return mediaPlayer.settings.mute; }
			set { mediaPlayer.settings.mute = value; }
		}
		public int 音量
		{
			get { return mediaPlayer.settings.volume; }
			set { mediaPlayer.settings.volume = value; }
		}
		public double 速度
		{
			get { return mediaPlayer.settings.rate; }
			set
			{
				try
				{
					mediaPlayer.settings.rate = value;
				}
				catch (COMException ex)
				{
					throw new ProduireException("現在の状況では設定できません。HRESULT:" + ex.ErrorCode.ToString("X"));
				}
			}
		}
		public bool エラー表示
		{
			get { return mediaPlayer.settings.enableErrorDialogs; }
			set { mediaPlayer.settings.enableErrorDialogs = value; }
		}
		[非表示]
		public string 状態
		{
			get { return mediaPlayer.status; }
		}
		[非表示]
		public WMPPlayState 再生状態
		{
			get { return mediaPlayer.playState; }
		}

		#endregion

		#region イベント手順

		private void SetEventHandling()
		{
			mediaPlayer.PlayStateChange += new _WMPOCXEvents_PlayStateChangeEventHandler(mediaPlayer_PlayStateChange);
			mediaPlayer.CurrentItemChange += new _WMPOCXEvents_CurrentItemChangeEventHandler(mediaPlayer_CurrentItemChange);
			mediaPlayer.MediaChange += new _WMPOCXEvents_MediaChangeEventHandler(mediaPlayer_MediaChange);
			mediaPlayer.StatusChange += new _WMPOCXEvents_StatusChangeEventHandler(mediaPlayer_StatusChange);
		}

		void mediaPlayer_PlayStateChange(int NewState)
		{
			if (再生状態が変更された != null) 再生状態が変更された(this, EventArgs.Empty);
		}

		void mediaPlayer_CurrentItemChange(object pdispMedia)
		{
			if (現在メディアが変更された != null)
			{
				media = new メディア情報(pdispMedia as IWMPMedia);
				現在メディアが変更された(this, EventArgs.Empty);
			}
		}

		void mediaPlayer_MediaChange(object Item)
		{
			if (メディア情報が変更された != null)
			{
				media = new メディア情報(Item as IWMPMedia);
				メディア情報が変更された(this, EventArgs.Empty);
			}
		}

		void mediaPlayer_StatusChange()
		{
			if (状態が変更された != null) 状態が変更された(this, EventArgs.Empty);
		}

		#endregion

		#region IDisposable メンバー

		public void Dispose()
		{
			mediaPlayer.controls.stop();
		}

		#endregion
	}

}
