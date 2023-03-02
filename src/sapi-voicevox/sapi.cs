using Microsoft.VisualBasic;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using TTSEngineLib;

namespace Yarukizero.Net.Sapi.VoiceVox;

internal static class GuidConst {
	public const string InterfaceGuid = "B3EBEE6A-4CF3-40AE-873E-6EB370FAC38D";
	public const string ClassGuid = "9868DCD8-1804-4280-94BD-E4AEB57B3BD2";
}

[ComVisible(true)]
[Guid(GuidConst.InterfaceGuid)]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IVoiceVoxTTSEngine : ISpTTSEngine, ISpObjectWithToken { }

[ComVisible(true)]
[Guid(GuidConst.ClassGuid)]
public class VoiceVoxTTSEngine : IVoiceVoxTTSEngine {
	private const ushort WAVE_FORMAT_PCM = 1;

	private static readonly Guid SPDFID_WaveFormatEx = new Guid("C31ADBAE-527F-4ff5-A230-F62BB61FF70C");
	private static readonly Guid SPDFID_Text = new Guid("7CEEF9F9-3D13-11d2-9EE7-00C04F797396");

	[DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
	private static extern IntPtr CreateFile(string pszFileName, int dwAccess, int dwShare, IntPtr psa, int dwCreatDisposition, int dwFlagsAndAttributes, IntPtr hTemplate);

	[DllImport("kernel32.dll")]
	private static extern bool ReadFile(IntPtr hFile, byte[] pBuffer, int nNumberOfBytesToRead, out int pNumberOfBytesRead, IntPtr pOverlapped);

	[DllImport("kernel32.dll")]
	private static extern int GetFileSize(IntPtr hFile, IntPtr pFileSizeHigh);
	[DllImport("kernel32.dll")]
	private static extern bool CloseHandle(IntPtr hObject);

	[DllImport("kernel32.dll")]
	private static extern int WaitForSingleObject(IntPtr hHandle, int dwMilliseconds);
	private const int WAIT_TIMEOUT = 0x102;
	private const int GENERIC_WRITE = 0x40000000;
	private const int GENERIC_READ = unchecked((int)0x80000000);
	private const int FILE_SHARE_READ = 0x00000001;
	private const int FILE_SHARE_WRITE = 0x00000002;
	private const int FILE_SHARE_DELETE = 0x00000004;
	private const int CREATE_NEW = 1;
	private const int CREATE_ALWAYS = 2;
	private const int OPEN_EXISTING = 3;
	private const int OPEN_ALWAYS = 4;
	private const int TRUNCATE_EXISTING = 5;

	[Flags]
	enum SPVESACTIONS {
		SPVES_CONTINUE = 0,
		SPVES_ABORT = (1 << 0),
		SPVES_SKIP = (1 << 1),
		SPVES_RATE = (1 << 2),
		SPVES_VOLUME = (1 << 3)
	}

	enum SPEVENTENUM {
		SPEI_UNDEFINED = 0,
		SPEI_START_INPUT_STREAM = 1,
		SPEI_END_INPUT_STREAM = 2,
		SPEI_VOICE_CHANGE = 3,
		SPEI_TTS_BOOKMARK = 4,
		SPEI_WORD_BOUNDARY = 5,
		SPEI_PHONEME = 6,
		SPEI_SENTENCE_BOUNDARY = 7,
		SPEI_VISEME = 8,
		SPEI_TTS_AUDIO_LEVEL = 9,
		SPEI_TTS_PRIVATE = 15,
		SPEI_MIN_TTS = 1,
		SPEI_MAX_TTS = 15,
		SPEI_END_SR_STREAM = 34,
		SPEI_SOUND_START = 35,
		SPEI_SOUND_END = 36,
		SPEI_PHRASE_START = 37,
		SPEI_RECOGNITION = 38,
		SPEI_HYPOTHESIS = 39,
		SPEI_SR_BOOKMARK = 40,
		SPEI_PROPERTY_NUM_CHANGE = 41,
		SPEI_PROPERTY_STRING_CHANGE = 42,
		SPEI_FALSE_RECOGNITION = 43,
		SPEI_INTERFERENCE = 44,
		SPEI_REQUEST_UI = 45,
		SPEI_RECO_STATE_CHANGE = 46,
		SPEI_ADAPTATION = 47,
		SPEI_START_SR_STREAM = 48,
		SPEI_RECO_OTHER_CONTEXT = 49,
		SPEI_SR_AUDIO_LEVEL = 50,
		SPEI_SR_RETAINEDAUDIO = 51,
		SPEI_SR_PRIVATE = 52,
		SPEI_ACTIVE_CATEGORY_CHANGED = 53,
		SPEI_RESERVED5 = 54,
		SPEI_RESERVED6 = 55,
		SPEI_MIN_SR = 34,
		SPEI_MAX_SR = 55,
		SPEI_RESERVED1 = 30,
		SPEI_RESERVED2 = 33,
		SPEI_RESERVED3 = 63
	}

	private const ulong SPFEI_FLAGCHECK = (1u << (int)SPEVENTENUM.SPEI_RESERVED1) | (1u << (int)SPEVENTENUM.SPEI_RESERVED2);
	private const ulong SPFEI_ALL_TTS_EVENTS = 0x000000000000FFFEul | SPFEI_FLAGCHECK;
	private const ulong SPFEI_ALL_SR_EVENTS = 0x003FFFFC00000000ul | SPFEI_FLAGCHECK;
	private const ulong SPFEI_ALL_EVENTS = 0xEFFFFFFFFFFFFFFFul;

	private ulong SPFEI(SPEVENTENUM SPEI_ord) => (1ul << (int)SPEI_ord) | SPFEI_FLAGCHECK;

	enum SPEVENTLPARAMTYPE {
		SPET_LPARAM_IS_UNDEFINED = 0,
		SPET_LPARAM_IS_TOKEN = (SPET_LPARAM_IS_UNDEFINED + 1),
		SPET_LPARAM_IS_OBJECT = (SPET_LPARAM_IS_TOKEN + 1),
		SPET_LPARAM_IS_POINTER = (SPET_LPARAM_IS_OBJECT + 1),
		SPET_LPARAM_IS_STRING = (SPET_LPARAM_IS_POINTER + 1)
	}

	private static readonly HttpClient httpClient = new HttpClient();

	private static readonly string KeyVoiceVoxEndPoint = "x-voicevox";
	private static readonly string KeyVoiceVoxSpeakerId = "x-voicevox-speaker";
	private static readonly string KeyVoiceVoxSpeakerPitch = "x-voicevox-speaker-pitch";
	private static readonly string KeyVoiceVoxSpeakerIntonation = "x-voicevox-speaker-intonation";
	private static readonly string KeyVoiceVoxSpeakerVolume = "x-voicevox-speaker-volume";

	private static readonly double DefaultVoicevoxSpeakerPitchScale = 0;
	private static readonly double DefaultVoicevoxSpeakerIntonationScale = 1;
	private static readonly double DefaultVoicevoxSpeakerVolumeScale = 1;



	private ISpObjectToken? token;
	private string voicevoxEndPoint = "http://127.0.0.1:50021";
	private string voicevoxSpeakerId = "1";
	private double voicevoxSpeakerPitchScale = DefaultVoicevoxSpeakerPitchScale;
	private double voicevoxSpeakerIntonationScale = DefaultVoicevoxSpeakerIntonationScale;
	private double voicevoxSpeakerVolumeScale = DefaultVoicevoxSpeakerVolumeScale;
	private int voicevoxSpeakerOutputSamplingRate = 44100;
	private bool voicevoxSpeakerOutputStereo = false;
	private System.Media.SoundPlayer? player = null;

	public void Speak(uint dwSpeakFlags, ref Guid rguidFormatId, ref WAVEFORMATEX pWaveFormatEx, ref SPVTEXTFRAG pTextFragList, ISpTTSEngineSite pOutputSite) {
		static uint output(ISpTTSEngineSite output, byte[] data) {
			var pWavData = IntPtr.Zero;
			try {
				if(data.Length == 0) {
					output.Write(pWavData, 0u, out var written);
					return written;
				} else {
					pWavData = Marshal.AllocCoTaskMem(data.Length);
					Marshal.Copy(data, 0, pWavData, data.Length);
					output.Write(pWavData, (uint)data.Length, out var written);
					return written;
				}
			}
			finally {
				if(pWavData != IntPtr.Zero) {
					Marshal.FreeCoTaskMem(pWavData);
				}
			}
		}
		void play(string resourceName) {
			if(this.player == null) {
				this.player = new System.Media.SoundPlayer();
			}
			player.Stream = typeof(VoiceVoxTTSEngine)
				.Assembly
				.GetManifestResourceStream(resourceName);
			player.Play();
		}

		if(rguidFormatId == SPDFID_Text) {
			return;
		}

		var optSpeed = 1d;
		{
			pOutputSite.GetRate(out var spd);
			optSpeed = Math.Max(Math.Min(1d, spd / 10d), -1d) + 1;
		}
		/*
		var volume = 1f;
		{
			pOutputSite.GetVolume(out var vol);
			volume = vol / 100f;
		}
		*/

		try {
			var writtenWavLength = 0UL;
			var currentTextList = pTextFragList;
			while(true) {
				if(currentTextList.State.eAction == SPVACTIONS.SPVA_ParseUnknownTag) {
					goto next;
				}
				var text = Regex.Replace(
					currentTextList.pTextStart,
					@"<.+?>",
					"",
					RegexOptions.IgnoreCase);
				if(string.IsNullOrWhiteSpace(text)) {
					goto next;
				}
				if(((SPVESACTIONS)pOutputSite.GetActions()).HasFlag(SPVESACTIONS.SPVES_ABORT)) {
					return;
				}
				AddEventToSAPI(pOutputSite, currentTextList.pTextStart, text, writtenWavLength);

				writtenWavLength += Speak(
					text,
					optSpeed,
					pOutputSite,
					play,
					output);

			next:
				if(currentTextList.pNext == IntPtr.Zero) {
					break;
				} else {
					currentTextList = Marshal.PtrToStructure<SPVTEXTFRAG>(currentTextList.pNext);
				}
			}
		}
		catch {
			play($"{typeof(VoiceVoxTTSEngine).Namespace}.Resources.unknown-error.wav");
			throw;
		}
	}


	private uint Speak(
		string text,
		double speed,
		ISpTTSEngineSite pOutputSite,
		Action<string> play,
		Func<ISpTTSEngineSite, byte[], uint> output) {

		try {
			var entry = $@"{voicevoxEndPoint}";
			Data.AudioQuery? json;
			{
				using var request = new HttpRequestMessage(
					HttpMethod.Post,
					new Uri($"{entry}/audio_query?text={HttpUtility.UrlEncode(text)}&speaker={this.voicevoxSpeakerId}"));
				using var response = httpClient.SendAsync(request);
				response.Wait();
				if(response.Result.StatusCode != HttpStatusCode.OK) {
					throw new InvalidOperationException();
				}

				var @string = response.Result.Content.ReadAsStringAsync();
				@string.Wait();
				json = JsonConvert.DeserializeObject<Data.AudioQuery>(@string.Result);
				if(json == null) {
					throw new InvalidOperationException();
				}
			}

			json.SpeedScale = speed;
			json.PitchScale = this.voicevoxSpeakerPitchScale;
			json.IntonationScale = this.voicevoxSpeakerIntonationScale;
			json.VolumeScale = this.voicevoxSpeakerVolumeScale;
			json.OutputSamplingRate = this.voicevoxSpeakerOutputSamplingRate;
			json.OutputStereo = this.voicevoxSpeakerOutputStereo;

			{
				using var request = new HttpRequestMessage(
					HttpMethod.Post,
					new Uri($"{entry}/synthesis?speaker={this.voicevoxSpeakerId}&enable_interrogative_upspeak={true}")) {

					Content = new StringContent(json.ToString(), Encoding.UTF8, @"application/json"),
				};
				using var response = httpClient.SendAsync(request);
				//using var response = httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);
				response.Wait();
				if(response.Result.StatusCode != HttpStatusCode.OK) {
					throw new InvalidOperationException();
				}
				var stream = response.Result.Content.ReadAsStreamAsync();
				stream.Wait();
				using var ms = new MemoryStream();
				{
					byte[] b = new byte[76800];
					int ret;
					while(0<(ret=stream.Result.Read(b, 0, b.Length))) {
						ms.Write(b, 0, ret);
					}
					stream.Result.Dispose();
				}
				ms.Position = 0;
				return output(pOutputSite, ms.ToArray());
			}
		}
		catch(AggregateException) { // VOICEVOXと通信できないので空を返す
			play($"{typeof(VoiceVoxTTSEngine).Namespace}.Resources.vv-notfound.wav");
			return output(pOutputSite, new byte[4]);
		}
		catch(Exception) {
			throw;
		}
	}


	private void AddEventToSAPI(ISpTTSEngineSite outputSite, string allText, string speakTargetText, ulong writtenWavLength) {
		outputSite.GetEventInterest(out var ev);
		var list = new List<SPEVENT>();
		var wParam = (uint)speakTargetText.Length;
		var lParam = allText.IndexOf(speakTargetText);
		if((ev & SPFEI(SPEVENTENUM.SPEI_SENTENCE_BOUNDARY)) == SPFEI(SPEVENTENUM.SPEI_SENTENCE_BOUNDARY)) {
			list.Add(new SPEVENT() {
				eEventId = (ushort)SPEVENTENUM.SPEI_SENTENCE_BOUNDARY,
				elParamType = (ushort)SPEVENTLPARAMTYPE.SPET_LPARAM_IS_UNDEFINED,
				wParam = wParam,
				lParam = lParam,
				ullAudioStreamOffset = writtenWavLength
			});
		}
		if((ev & SPFEI(SPEVENTENUM.SPEI_WORD_BOUNDARY)) == SPFEI(SPEVENTENUM.SPEI_WORD_BOUNDARY)) {
			list.Add(new SPEVENT() {
				eEventId = (ushort)SPEVENTENUM.SPEI_WORD_BOUNDARY,
				elParamType = (ushort)SPEVENTLPARAMTYPE.SPET_LPARAM_IS_UNDEFINED,
				wParam = wParam,
				lParam = lParam,
				ullAudioStreamOffset = writtenWavLength
			});
		}
		if(list.Any()) {
			var arr = list.ToArray();
			outputSite.AddEvents(ref arr[0], (uint)arr.Length);
		}
	}

	public void GetOutputFormat(ref Guid pTargetFmtId, ref WAVEFORMATEX pTargetWaveFormatEx, out Guid pOutputFormatId, IntPtr ppCoMemOutputWaveFormatEx) {
		pOutputFormatId = SPDFID_WaveFormatEx;
		var wf = new WAVEFORMATEX() {
			wFormatTag = WAVE_FORMAT_PCM,
			nChannels = 1,
			cbSize = 0,
			nSamplesPerSec = 44100,
			wBitsPerSample = 16,
			nBlockAlign = 1 * 16 / 8, // チャンネル * bps / 8
			nAvgBytesPerSec = 44100 * (1 * 16 / 8), // サンプリングレート / ブロックアライン
		};

		var p = Marshal.AllocCoTaskMem(Marshal.SizeOf(wf));
		Marshal.StructureToPtr(wf, p, false);
		Marshal.WriteIntPtr(ppCoMemOutputWaveFormatEx, p);
	}

	public void SetObjectToken(ISpObjectToken pToken) {
		string get(string key) {
			try {
				pToken.GetStringValue(key, out var s);
				return s;
			}
			catch(COMException) {
				return "";
			}
		}
		double @double(string val, double @default) {
			try {
				return double.Parse(val);
			}
			catch {
				return @default;
			}
		}

		this.token = pToken;
		this.voicevoxEndPoint = get(KeyVoiceVoxEndPoint);
		this.voicevoxSpeakerId = get(KeyVoiceVoxSpeakerId);
		this.voicevoxSpeakerPitchScale = @double(get(KeyVoiceVoxSpeakerPitch), DefaultVoicevoxSpeakerPitchScale);
		this.voicevoxSpeakerIntonationScale = @double(get(KeyVoiceVoxSpeakerIntonation), DefaultVoicevoxSpeakerIntonationScale);
		this.voicevoxSpeakerVolumeScale = @double(get(KeyVoiceVoxSpeakerVolume), DefaultVoicevoxSpeakerVolumeScale);
	}


	public void GetObjectToken(ref ISpObjectToken? ppToken) {
		ppToken = token;
	}

	[ComRegisterFunction()]
	public static void RegisterClass(string _) {
		static string safePath(string name) => Regex.Replace(name, @"[\s,/\:\*\?""\<\>\|]", "");
		var entry = @"SOFTWARE\Microsoft\Speech\Voices\Tokens";
		var prefix = "TTS_YARUKIZERO_VOICEVOX";


		// 一度情報を破棄する
		InitializeRegistry();
		foreach(var it in new[] {
			new {
				Id = 2,
				Name = "ShikokuMethane",
				NameKana = "四国めたん(ノーマル)",
			},
			new {
				Id = 3,
				Name = "Zundamon",
				NameKana = "ずんだもん(ノーマル)",
			},
			new {
				Id = 8,
				Name = "KasukabeTsumugi",
				NameKana = "春日部つむぎ(ノーマル)",
			},
		}) {
			using(var registryKey = Registry.LocalMachine.CreateSubKey($@"{entry}\{prefix}-{safePath(it.Name)}")) {
				registryKey.SetValue("", $"VOICEVOX {it.NameKana}");
				registryKey.SetValue("411", $"VOICEVOX {it.NameKana}");
				registryKey.SetValue("CLSID", $"{{{GuidConst.ClassGuid}}}");
				registryKey.SetValue(KeyVoiceVoxEndPoint, "http://127.0.0.1:50021");
				registryKey.SetValue(KeyVoiceVoxSpeakerId, $"{it.Id}");
				registryKey.SetValue(KeyVoiceVoxSpeakerPitch, $"{DefaultVoicevoxSpeakerPitchScale:F2}");
				registryKey.SetValue(KeyVoiceVoxSpeakerIntonation, $"{DefaultVoicevoxSpeakerIntonationScale:F2}");
				registryKey.SetValue(KeyVoiceVoxSpeakerVolume, $"{DefaultVoicevoxSpeakerVolumeScale:F2}");
			}
			using(var registryKey = Registry.LocalMachine.CreateSubKey($@"{entry}\{prefix}-{safePath(it.Name)}\Attributes")) {
				registryKey.SetValue("Age", "Teen"); // ここはてきとー
				registryKey.SetValue("Vendor", "Hiroshiba Kazuyuki");
				registryKey.SetValue("Language", "411");
				registryKey.SetValue("Gender", "Female"); // ここもてきとー
				registryKey.SetValue("Name", $"VOICEVOX {it.Name}");
			}
		}
	}

	[ComUnregisterFunction()]
	public static void UnregisterClass(string _) {
		InitializeRegistry();
	}

	private static void InitializeRegistry() {
		using(var regTokensKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Speech\Voices\Tokens\", true)) {
			if(regTokensKey == null) {
				return;
			}
			foreach(var name in regTokensKey.GetSubKeyNames()) {
				using(var regKey = regTokensKey.OpenSubKey(name)) {
					if(regKey?.GetValue("CLSID") is string id && id == $"{{{GuidConst.ClassGuid}}}") {
						regTokensKey.DeleteSubKeyTree(name);
					}
				}
			}
		}
	}
}