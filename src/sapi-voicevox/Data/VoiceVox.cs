using Microsoft.SqlServer.Server;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Net.Mime.MediaTypeNames;

namespace Yarukizero.Net.Sapi.VoiceVox.Data {
	public class JsonObject {
		public override string ToString() => JsonConvert.SerializeObject(this);
	}

	public class AudioQuery : JsonObject {
		[JsonProperty("accent_phrases", Required = Required.Always)]
		public AccentPhrases[] AccentPhrases { get; set; }

		[JsonProperty("speedScale", Required = Required.Always)]
		public double SpeedScale { get; set; }

		[JsonProperty("pitchScale", Required = Required.Always)]
		public double PitchScale { get; set; }

		[JsonProperty("intonationScale", Required = Required.Always)]
		public double IntonationScale { get; set; }

		[JsonProperty("volumeScale", Required = Required.Always)]
		public double VolumeScale { get; set; }

		[JsonProperty("prePhonemeLength", Required = Required.Always)]
		public double PrePhonemeLength { get; set; }


		[JsonProperty("postPhonemeLength", Required = Required.Always)]
		public double PostPhonemeLength { get; set; }


		[JsonProperty("outputSamplingRate", Required = Required.Always)]
		public int OutputSamplingRate { get; set; }

		[JsonProperty("outputStereo", Required = Required.Always)]
		public bool OutputStereo { get; set; }

		[JsonProperty("kana")]
		public string? Kana { get; set; }
	}

	public class AccentPhrases : JsonObject {
		[JsonProperty("moras", Required = Required.Always)]
		public Mora[] Moras { get; set; }

		[JsonProperty("accent", Required = Required.Always)]
		public int Accent { get; set; }

		[JsonProperty("pause_mora")]
		public PauseMora? PauseMora { get; set; }

		[JsonProperty("is_interrogative")]
		public bool? IsInterrogative { get; set; }
	}


	public class Mora : JsonObject {
		[JsonProperty("text", Required = Required.Always)]
		public string Text { get; set; }

		[JsonProperty("consonant")]
		public string? Consonant { get; set; }

		[JsonProperty("consonant_length")]
		public double? consonant_length { get; set; }

		[JsonProperty("vowel", Required = Required.Always)]
		public string Vowel { get; set; }

		[JsonProperty("vowel_length", Required = Required.Always)]
		public double VowelLength { get; set; }

		[JsonProperty("pitch", Required = Required.Always)]
		public double Pitch { get; set; }
	}

	public class PauseMora : JsonObject {
		[JsonProperty("text", Required = Required.Always)]
		public string Text { get; set; }

		[JsonProperty("consonant")]
		public string? Consonant { get; set; }

		[JsonProperty("consonant_length")]
		public double? ConsonantLength { get; set; }

		[JsonProperty("vowel", Required = Required.Always)]
		public string Vowel { get; set; }

		[JsonProperty("vowel_length", Required = Required.Always)]
		public double VowelLength { get; set; }

		[JsonProperty("pitch", Required = Required.Always)]
		public double Pitch { get; set; }
	}

	public class Speaker : JsonObject {
		[JsonProperty("supported_features")]
		public SpeakerSupportedFeature? SupportedFeatures { get; set; }

		[JsonProperty("name", Required = Required.Always)]
		public string Name { get; set; }

		[JsonProperty("speaker_uuid", Required = Required.Always)]
		public string SpeakerUuid { get; set; }

		[JsonProperty("styles", Required = Required.Always)]
		public SpeakerStyle[] Styles { get; set; }

		[JsonProperty("version")]
		public string? Version { get; set; }
	}

	public class SpeakerSupportedFeature : JsonObject {
		[JsonProperty("permitted_synthesis_morphing")]
		public string? PermittedSynthesisMorphing { get; set; } // "ALL" "SELF_ONLY" "NOTHING"
	}

	public class SpeakerStyle : JsonObject {
		[JsonProperty("name", Required = Required.Always)]
		public string Name { get; set; }

		[JsonProperty("id", Required = Required.Always)]
		public int Id { get; set; }
	}
}