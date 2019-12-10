using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Produire;
using AxWMPLib;
using System.Reflection;
using System.IO;
using WMPLib;
using System.Drawing;
using System.Runtime.InteropServices;

[assembly: ResourceAssembly("Produire.WMP.Interop.WMPLib.dll")]
[assembly: ResourceAssembly("Produire.WMP.AxInterop.WMPLib.dll")]

namespace Produire.WMP
{
	[種類(DocUrl = "https://rdr.utopiat.net/docs/plugins/RdrWMP/wmpctrl.htm"), メインスレッド]
	public class WindowsMediaPlayer部品 : AxWindowsMediaPlayer, IProduireClass
	{
		//http://msdn.microsoft.com/en-us/library/aa472935.aspx
		メディア情報 media;
		bool initialized = false;

		#region イベント手順

		public event EventHandler 状態が変更された;
		public event EventHandler 再生状態が変更された;
		public event EventHandler 現在メディアが変更された;
		public event EventHandler メディア情報が変更された;
		public event ProduireEventHandler 再生位置が変更された;

		#endregion

		public WindowsMediaPlayer部品()
		{

		}

		#region 手順

		[自分("で")]
		public void 開く([を]string アドレス)
		{
			if (!initialized) 初期化();
			string lurl = アドレス.ToLower();
			if (lurl.StartsWith("http:") || lurl.StartsWith("https:"))
			{
				try
				{
					launchURL(アドレス);
				}
				catch (COMException ex)
				{
					throw new ProduireException("アドレスを開くことができません。HRESULT:" + ex.ErrorCode.ToString("X"));
				}
			}
			else
			{
				try
				{
					this.URL = アドレス;
				}
				catch (COMException ex)
				{
					throw new ProduireException("ファイルを開くことができません。HRESULT:" + ex.ErrorCode.ToString("X"));
				}
			}
		}

		[自分("を")]
		public void 再生()
		{
			if (!initialized) 初期化();
			try
			{
				Ctlcontrols.play();
			}
			catch (COMException ex)
			{
				throw new ProduireException("再生できない状態です。HRESULT:" + ex.ErrorCode.ToString("X"));
			}
		}
		[自分("を")]
		public void 一時停止()
		{
			if (!initialized) 初期化();
			try
			{
				Ctlcontrols.pause();
			}
			catch (COMException ex)
			{
				throw new ProduireException("一時停止が利用できない状態です。HRESULT:" + ex.ErrorCode.ToString("X"));
			}
		}
		[自分("を")]
		public void 停止()
		{
			if (!initialized) 初期化();
			try
			{
				Ctlcontrols.stop();
			}
			catch (COMException ex)
			{
				throw new ProduireException("停止が利用できない状態です。HRESULT:" + ex.ErrorCode.ToString("X"));
			}
		}
		[自分("を")]
		public void 次へ進む()
		{
			if (!initialized) 初期化();
			try
			{
				Ctlcontrols.next();
			}
			catch (COMException ex)
			{
				throw new ProduireException("この操作が利用できない状態です。HRESULT:" + ex.ErrorCode.ToString("X"));
			}
		}
		[自分("を")]
		public void 前へ戻る()
		{
			if (!initialized) 初期化();
			try
			{
				Ctlcontrols.previous();
			}
			catch (COMException ex)
			{
				throw new ProduireException("この操作が利用できない状態です。HRESULT:" + ex.ErrorCode.ToString("X"));
			}
		}
		[自分("を")]
		public void 初期化()
		{
			SetEventHandling();
			EndInit();
			initialized = true;
		}
		[自分("を")]
		public void 初期化開始()
		{
			初期化();
		}
		#endregion

		#region プロパティ

		[非表示]
		public string アドレス
		{
			get { return base.URL; }
			set
			{
				try
				{
					base.URL = value;
				}
				catch (COMException ex)
				{
					throw new ProduireException("アドレスを開くことができません。HRESULT:" + ex.ErrorCode.ToString("X"));
				}
			}
		}
		[非表示]
		public double 再生位置
		{
			get { return Ctlcontrols.currentPosition; }
			set
			{
				try
				{
					Ctlcontrols.currentPosition = value;
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
			get { return Ctlcontrols.currentPositionString; }
		}
		[非表示]
		public メディア情報 現在メディア
		{
			get { return media; }
			set
			{
				try
				{
					currentMedia = value.media;
				}
				catch (COMException ex)
				{
					throw new ProduireException("現在の再生状態では設定できません。HRESULT:" + ex.ErrorCode.ToString("X"));
				}
			}
		}
		public bool コンテキストメニュー
		{
			get
			{
				if (!initialized) 初期化();
				return base.enableContextMenu;
			}
			set
			{
				if (!initialized) 初期化();
				base.enableContextMenu = value;
			}
		}
		[非表示(VisibilityTarget.FormDesigner)]
		public bool フルスクリーン
		{
			get
			{
				if (!initialized) 初期化();
				return base.fullScreen;
			}
			set
			{
				if (!initialized) 初期化();
				try
				{
					base.fullScreen = value;
				}
				catch (COMException ex)
				{
					throw new ProduireException("現在の再生状態では設定できません。HRESULT:" + ex.ErrorCode.ToString("X"));
				}
			}
		}
		public bool 自動再生
		{
			get { return settings.autoStart; }
			set { settings.autoStart = value; }
		}
		public int バランス
		{
			get { return settings.balance; }
			set { settings.balance = value; }
		}
		public bool 消音
		{
			get { return settings.mute; }
			set { settings.mute = value; }
		}
		public int 音量
		{
			get { return settings.volume; }
			set { settings.volume = value; }
		}
		public double 速度
		{
			get { return settings.rate; }
			set
			{
				try
				{
					settings.rate = value;
				}
				catch (COMException ex)
				{
					throw new ProduireException("現在の状況では設定できません。HRESULT:" + ex.ErrorCode.ToString("X"));
				}
			}
		}
		public bool エラー表示
		{
			get { return settings.enableErrorDialogs; }
			set { settings.enableErrorDialogs = value; }
		}
		[非表示]
		public string 状態
		{
			get { return base.status; }
		}
		[非表示]
		public WMPPlayState 再生状態
		{
			get { return base.playState; }
		}
		public 画面状態 画面状態
		{
			get
			{
				switch (uiMode)
				{
					case "none":
						return 画面状態.なし;
					case "mini":
						return 画面状態.最小;
					case "full":
						return 画面状態.フル;
					default:
						return 画面状態.カスタム;
				}
			}
			set
			{
				switch (value)
				{
					case 画面状態.なし:
						base.uiMode = "none";
						break;
					case 画面状態.最小:
						base.uiMode = "mini";
						break;
					case 画面状態.フル:
						base.uiMode = "full";
						break;
					case 画面状態.カスタム:
						base.uiMode = "custom";
						break;
				}
			}
		}

		#endregion

		#region イベント手順

		private void SetEventHandling()
		{
			base.PlayStateChange += new AxWMPLib._WMPOCXEvents_PlayStateChangeEventHandler(WindowsMediaPlayer部品_PlayStateChange);
			base.CurrentItemChange += new AxWMPLib._WMPOCXEvents_CurrentItemChangeEventHandler(WindowsMediaPlayer部品_CurrentItemChange);
			base.PositionChange += new AxWMPLib._WMPOCXEvents_PositionChangeEventHandler(WindowsMediaPlayer部品_PositionChange);
			base.MediaChange += new AxWMPLib._WMPOCXEvents_MediaChangeEventHandler(WindowsMediaPlayer部品_MediaChange);
			base.StatusChange += new EventHandler(WindowsMediaPlayer部品_StatusChange);
		}

		void WindowsMediaPlayer部品_PlayStateChange(object sender, _WMPOCXEvents_PlayStateChangeEvent e)
		{
			if (再生状態が変更された != null) 再生状態が変更された(this, EventArgs.Empty);
		}

		void WindowsMediaPlayer部品_StatusChange(object sender, EventArgs e)
		{
			if (状態が変更された != null) 状態が変更された(this, EventArgs.Empty);
		}

		void WindowsMediaPlayer部品_MediaChange(object sender, _WMPOCXEvents_MediaChangeEvent e)
		{
			if (メディア情報が変更された != null)
			{
				media = new メディア情報(e.item as IWMPMedia);
				メディア情報が変更された(this, EventArgs.Empty);
			}
		}

		void WindowsMediaPlayer部品_PositionChange(object sender, _WMPOCXEvents_PositionChangeEvent e)
		{
			if (再生位置が変更された != null) 再生位置が変更された(this, new 再生位置イベント情報(e));
		}

		void WindowsMediaPlayer部品_CurrentItemChange(object sender, _WMPOCXEvents_CurrentItemChangeEvent e)
		{
			if (現在メディアが変更された != null)
			{
				media = new メディア情報(e.pdispMedia as IWMPMedia);
				現在メディアが変更された(this, EventArgs.Empty);
			}
		}

		#endregion
	}

	[列挙体(typeof(WMPPlayState))]
	public enum 再生状態
	{
		再生中 = WMPPlayState.wmppsPlaying,
		停止中 = WMPPlayState.wmppsStopped,
		バッファリング中 = WMPPlayState.wmppsBuffering,
		一時停止中 = WMPPlayState.wmppsPaused,
		準備完了 = WMPPlayState.wmppsReady,
		再生終了 = WMPPlayState.wmppsMediaEnded,
		再接続中 = WMPPlayState.wmppsReconnecting,
		待機中 = WMPPlayState.wmppsWaiting,
		再生準備中 = WMPPlayState.wmppsTransitioning,
		最後 = WMPPlayState.wmppsLast
	}

	[列挙体]
	public enum 画面状態
	{
		なし, 最小, フル, カスタム
	}

	[種類(DocUrl = "https://rdr.utopiat.net/docs/plugins/RdrWMP/mediainfo.htm")]
	public class メディア情報 : IProduireClass
	{
		internal IWMPMedia media;

		public メディア情報(IWMPMedia media)
		{
			this.media = media;
		}

		#region 手順

		[自分("から"), 補語("属性を"), 手順("取得")]
		public string 属性を取得([という]string 属性名)
		{
			return media.getItemInfo(属性名);
		}

		#endregion

		#region 設定項目

		public string アドレス
		{
			get { return media.sourceURL; }
		}

		public string 名前
		{
			get { return media.name; }
		}

		public double 長さ
		{
			get { return media.duration; }
		}

		public string 長さ情報
		{
			get { return media.durationString; }
		}

		public Size 大きさ
		{
			get { return new Size(media.imageSourceWidth, media.imageSourceHeight); }
		}

		public int 幅
		{
			get { return media.imageSourceWidth; }
		}

		public int 高さ
		{
			get { return media.imageSourceHeight; }
		}

		public Dictionary<string, string> 属性一覧
		{
			get
			{
				int count = media.attributeCount;
				Dictionary<string, string> result = new Dictionary<string, string>();
				for (int i = 0; i < count; i++)
				{
					string name = media.getAttributeName(i);
					string value = media.getItemInfo(name);
					result[name] = value;
				}
				return result;
			}
		}

		#endregion
	}

	public class 再生位置イベント情報 : ProduireEventArgs, IProduireClass
	{
		double oldPosition;
		double newPosition;

		public 再生位置イベント情報(_WMPOCXEvents_PositionChangeEvent e)
		{
			oldPosition = e.oldPosition;
			newPosition = e.newPosition;
		}

		public 再生位置イベント情報(double oldPosition, double newPosition)
		{
			this.oldPosition = oldPosition;
			this.newPosition = newPosition;
		}

		#region 設定項目

		public double 旧位置
		{
			get { return oldPosition; }
		}

		public double 位置
		{
			get { return newPosition; }
		}

		#endregion
	}

	public class メディア情報イベント情報 : ProduireEventArgs, IProduireClass
	{
	}
}
