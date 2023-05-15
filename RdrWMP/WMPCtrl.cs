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
		public event EventHandler エラーが発生した;

		#endregion

		public WindowsMediaPlayer部品()
		{

		}

		#region 手順

		[自分("で")]
		public void 開く([を] string アドレス)
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
		[自分("を"), 補語("次へ"), 動詞("進む")]
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
		[自分("を"), 補語("前へ"), 動詞("戻る")]
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
			base.PlayStateChange += WMP_PlayStateChange;
			base.CurrentItemChange += WMP_CurrentItemChange;
			base.PositionChange += WMP_PositionChange;
			base.MediaChange += WMP_MediaChange;
			base.StatusChange += WMP_StatusChange;

			base.ClickEvent += WMP_ClickEvent;
			base.DoubleClickEvent += WMP_DoubleClickEvent;

			base.KeyDownEvent += WMP_KeyDownEvent;
			base.KeyPressEvent += WMP_KeyPressEvent;
			base.KeyUpEvent += WMP_KeyUpEvent;
			base.MouseDownEvent += WMP_MouseDownEvent;
			base.MouseMoveEvent += WMP_MouseMoveEvent;
			base.MouseUpEvent += WMP_MouseUpEvent;
			base.ErrorEvent += WMP_ErrorEvent;
		}

		void WMP_PlayStateChange(object sender, _WMPOCXEvents_PlayStateChangeEvent e)
		{
			if (再生状態が変更された != null) 再生状態が変更された(this, EventArgs.Empty);
		}

		void WMP_StatusChange(object sender, EventArgs e)
		{
			if (状態が変更された != null) 状態が変更された(this, EventArgs.Empty);
		}

		void WMP_MediaChange(object sender, _WMPOCXEvents_MediaChangeEvent e)
		{
			if (メディア情報が変更された != null)
			{
				media = new メディア情報(e.item as IWMPMedia);
				メディア情報が変更された(this, EventArgs.Empty);
			}
		}

		void WMP_PositionChange(object sender, _WMPOCXEvents_PositionChangeEvent e)
		{
			if (再生位置が変更された != null) 再生位置が変更された(this, new 再生位置イベント情報(e));
		}

		void WMP_CurrentItemChange(object sender, _WMPOCXEvents_CurrentItemChangeEvent e)
		{
			if (現在メディアが変更された != null)
			{
				media = new メディア情報(e.pdispMedia as IWMPMedia);
				現在メディアが変更された(this, EventArgs.Empty);
			}
		}

		private void WMP_ErrorEvent(object sender, EventArgs e)
		{
			if (エラーが発生した != null) エラーが発生した(this, EventArgs.Empty);
		}

		#endregion

		#region キーボード

		private void WMP_KeyDownEvent(object sender, _WMPOCXEvents_KeyDownEvent e)
		{
			var e2 = new KeyEventArgs((Keys)e.nKeyCode);
			base.OnKeyDown(e2);
		}

		protected override void OnKeyDown(KeyEventArgs e)
		{
		}

		private void WMP_KeyPressEvent(object sender, _WMPOCXEvents_KeyPressEvent e)
		{
			base.OnKeyPress(new KeyPressEventArgs((char)e.nKeyAscii));
		}
		protected override void OnKeyPress(KeyPressEventArgs e)
		{
		}

		private void WMP_KeyUpEvent(object sender, _WMPOCXEvents_KeyUpEvent e)
		{
			var e2 = new KeyEventArgs((Keys)e.nKeyCode);
			base.OnKeyUp(e2);
		}
		protected override void OnKeyUp(KeyEventArgs e)
		{
		}

		#endregion

		#region マウス

		private void WMP_MouseDownEvent(object sender, _WMPOCXEvents_MouseDownEvent e)
		{
			base.OnMouseDown(new MouseEventArgs(GetButtons(e.nButton), 1, e.fX, e.fY, 0));
		}
		protected override void OnMouseDown(MouseEventArgs e)
		{
		}

		private void WMP_MouseMoveEvent(object sender, _WMPOCXEvents_MouseMoveEvent e)
		{
			base.OnMouseMove(new MouseEventArgs(GetButtons(e.nButton), 1, e.fX, e.fY, 0));
		}
		protected override void OnMouseMove(MouseEventArgs e)
		{
		}

		private void WMP_MouseUpEvent(object sender, _WMPOCXEvents_MouseUpEvent e)
		{
			base.OnMouseUp(new MouseEventArgs(GetButtons(e.nButton), 1, e.fX, e.fY, 0));
		}
		protected override void OnMouseUp(MouseEventArgs e)
		{
		}

		/// <summary></summary>
		public MouseButtons GetButtons(short nButton)
		{
			if ((nButton & 1) != 0)
				return MouseButtons.Left;
			else if ((nButton & 2) != 0)
				return MouseButtons.Right;
			else if ((nButton & 4) != 0)
				return MouseButtons.Middle;
			else
				return MouseButtons.None;
		}


		private void WMP_ClickEvent(object sender, _WMPOCXEvents_ClickEvent e)
		{
			base.OnClick(EventArgs.Empty);
		}

		private void WMP_DoubleClickEvent(object sender, _WMPOCXEvents_DoubleClickEvent e)
		{
			base.OnDoubleClick(EventArgs.Empty);
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
		public string 属性を取得([という] string 属性名)
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
