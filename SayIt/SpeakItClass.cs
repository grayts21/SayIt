using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Collections;
using System.Text;
using System.Threading;
using System.Xml;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;

namespace SpeakIt
{
	class SpeakItClass
	{
		SpeechLib.SpVoice voice;
		Dictionary<String, String> WordReplace;
		Dictionary<String, bool> MethodFlag;
		Dictionary<String, String> MethodReplace;
		Dictionary<String, String> MethodRegEx;
		Dictionary<String, String> CharReplace;

		class SpeakAttr
		{
			public string _pre;
			public string _preAny;
			public string _post;
			public string _postAny;
			public string _Token;
			public string _ReplacePre;
			public string _ReplacePost;
			public string _ReplaceToken;
			public bool _caseSensitive;

			public SpeakAttr (string pre,
							 string preAny,
							 string post,
							 string postAny,
							 string token,
							 string replacePre,
							 string replacePost,
							 string replaceToken,
							 bool caseSensitive)
			{
				_pre = pre;
				_preAny = preAny;
				_post = post;
				_postAny = postAny;
				_Token = token;
				_ReplacePre = replacePre;
				_ReplacePost = replacePost;
				_ReplaceToken = replaceToken;
				_caseSensitive = caseSensitive;
			}

			public SpeakAttr ()
			{
				_pre = "~";
				_preAny = "~";
				_post = "~";
				_postAny = "~";
				_Token = "~";
				_ReplacePre = "~";
				_ReplacePost = "~";
				_ReplaceToken = "~";
				_caseSensitive = false;
			}

			public override string ToString ()
			{
				return _pre + "}{" + _preAny + "}{" + _Token + "}{" + _post + "}{" + _postAny + "}{" + _ReplacePre + "}{" + _ReplaceToken + "}{" + _ReplacePost + "}{" + _caseSensitive.ToString() + "}";
			}
		}

		List<SpeakAttr> SpeakCmds;

		SpeechLib.SpeechVoiceSpeakFlags speakFlag;
		string RuleName = string.Empty;
		bool writeToFile = false;
		string outputFilename = string.Empty;
		const string wordCharDel = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890";

		class PrePost
		{
			public PrePost (string s)
			{
				_s = s;
			}
			private string _s;
			public string s
			{
				get { return _s; }
				set { _s = value; }
			}

			public override string ToString ()
			{
				return _s;
			}
		}

		class wordToken
		{

			public PrePost PreToken;
			public string Token;
			public PrePost PostToken;

			public wordToken (PrePost _preToken, string _token, PrePost _postToken)
			{
				PreToken = _preToken;
				Token = _token;
				PostToken = _postToken;
			}

			public override string ToString ()
			{
				return PreToken.s + "}{" + Token + "}{" + PostToken.s;
			}
		}

		LinkedList<wordToken> tokens = new LinkedList<wordToken>();

		/// <summary>Short Summary of Program.InitSpeech</summary>
		private bool InitSpeech ()
		{
			try
			{
				voice = new SpeechLib.SpVoice();

				SpeechLib.ISpeechObjectTokens tokens = voice.GetVoices("", "");
				voice.Voice = tokens.Item(0);
				SpeechLib.ISpeechObjectTokens Atokens = voice.GetAudioOutputs(null, "");
				voice.AudioOutput = Atokens.Item(0);

				voice.AudioOutputStream.Format.Type = SpeechLib.SpeechAudioFormatType.SAFT22kHz16BitMono;
				voice.AudioOutputStream = voice.AudioOutputStream;
				voice.Rate = 8;
				voice.AllowAudioOutputFormatChangesOnNextSet = false;
				//voice.EventInterests = SpeechLib.SpeechVoiceEvents.SVESentenceBoundary;
				//voice.Sentence += new SpeechLib._ISpeechVoiceEvents_SentenceEventHandler(sentence_Event);

				speakFlag = SpeechLib.SpeechVoiceSpeakFlags.SVSFDefault
					| SpeechLib.SpeechVoiceSpeakFlags.SVSFlagsAsync
					| SpeechLib.SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak
					//| SpeechLib.SpeechVoiceSpeakFlags.SVSFNLPSpeakPunc
					//| SpeechLib.SpeechVoiceSpeakFlags.SVSFIsXML;
					;
				//int rate = voice.Rate;
				//SpeechLib.SpObjectToken audio = voice.AudioOutput;
				//int volume = voice.Volume;
				//SpeechLib.SpObjectToken thevoice = voice.Voice;
				//SpeechLib.ISpeechObjectToken token;
				return true;

			}
			catch (Exception)
			{
				return false;
			}
		}

		/// <summary>Short Summary of Program.GetTextToSpeak</summary>
		/// <param name="getFrom">Method to obtain text, clipboard, selection, file</param>
		/// <param name="outTextToSpeak">The Text to speak as a string</param>
		private bool GetTextToSpeak (int getFrom, out string outTextToSpeak)
		{
			try
			{
				if (getFrom == 100)
				{
					outTextToSpeak = "KeyValuePair \"Tony\", YouAre a c# Programer. You uses .net and metadata to 'make' programs customizable with <xml>.";
					return true;
				}
				if (getFrom == 1)
				{
					outTextToSpeak = "I could not find anything in the clipboard to tell you!";
					if (Clipboard.ContainsText())
					{
						outTextToSpeak = Clipboard.GetText();
					}

					return true;
				}
				else
				{
					outTextToSpeak = "Nothing to say!";
					return false;
				}
			}
			catch (Exception)
			{
				throw;
			}

			//return true;
		}

		/// <summary>Process_MultiWordCaps: Break words with Caps to multiple words</summary>
		/// <param name="regEx">Regular Expression to use</param>
		/// <param name="replace">Text to use to replace</param>
		private void Process_MultiWordCaps (string regEx, string replace)
		{
			//StringBuilder toReplace = new StringBuilder(50);
			List<string> replaceList = new List<string>(50);

			StringBuilder currWord = new StringBuilder();
			StringBuilder upperWord = new StringBuilder();

			bool foundMulti1 = false;
			bool foundMulti2 = false;
			bool foundMulti3 = false;

			try
			{
				//LinkedListNode<wordToken> item = tokens.First;
				for (LinkedListNode<wordToken> item = tokens.First;
					item != null;
					item = item.Next)
				{
					currWord.Length = 0;
					upperWord.Length = 0;

					foundMulti1 = false;
					foundMulti2 = false;
					foundMulti3 = false;

					foreach (char wChar in item.Value.Token)
					{
						if ((wChar.CompareTo('A') >= 0) && (wChar.CompareTo('Z') <= 0))
						{
							// Process Upper case letter
							if (foundMulti1)
							{
								foundMulti2 = true;
							}
							foundMulti1 = true;
							if (currWord.Length == 1)
							{
								upperWord.Append(currWord);
								currWord.Length = 0;
							}
							else if (currWord.Length > 1)
							{
								if (upperWord.Length > 0)
								{
									replaceList.Add(upperWord.ToString());
									upperWord.Length = 0;
								}
								replaceList.Add(currWord.ToString());
								currWord.Length = 0;
							}
							currWord.Append(wChar);
						}
						else // Process lower case letter
						{
							foundMulti3 = true;
							currWord.Append(wChar);
						}
					}

					if (currWord.Length == 1)
					{
						if ((currWord[0].CompareTo('A') >= 0) && (currWord[0].CompareTo('Z') <= 0))
						{
							upperWord.Append(currWord);
							currWord.Length = 0;
						}
					}

					if (upperWord.Length > 0)
					{
						replaceList.Add(upperWord.ToString());
					}

					if (currWord.Length > 0)
					{
						replaceList.Add(currWord.ToString());
					}
					currWord.Length = 0;
					upperWord.Length = 0;

					// Check if word contained Multi-Words
					if (foundMulti1 && foundMulti2 && foundMulti3)
					{
						PrePost spacePrePost = new PrePost(replace);
						wordToken newToken = new wordToken(item.Value.PreToken, replaceList[0], spacePrePost);
						tokens.AddBefore(item, newToken);

						for (int i = 1; i < replaceList.Count; i++)
						{
							PrePost AnotherPrePost = new PrePost(replace);
							newToken = new wordToken(spacePrePost, replaceList[i], AnotherPrePost);
							tokens.AddBefore(item, newToken);
						}
						newToken.PostToken = item.Value.PostToken;
						LinkedListNode<wordToken> PrevItem = item.Previous;
						tokens.Remove(item);
						item = PrevItem;
						PrevItem = null;
					}
					replaceList.Clear();
				}
			}
			catch (Exception)
			{
				throw;
			}
		}

		private String TickleWord (String word)
		{
			if (WordReplace.ContainsKey(word))
			{
				return WordReplace[word];
			}
			return word;
		}

		/// <summary>Process_Replace</summary>
		/// <param name="InTextToSpeak">Input string to process</param>
		private void Process_Replace ()
		{
			int startEdit = 0;
			int lengthEdit = 0;
			StringComparison currentComparison = StringComparison.CurrentCultureIgnoreCase;

			// Process each token against the "Speak" commands in WordReplace
			foreach (wordToken item in tokens)
			{
				// Process the attributes, Pre=, PreAny=, Post=, PostAny=, Token=
				// Replace with the attributes, ReplacePre=, ReplacePost=, ReplaceToken=
				foreach (SpeakAttr attr in SpeakCmds)
				{
					currentComparison = attr._caseSensitive ? StringComparison.CurrentCulture : StringComparison.CurrentCultureIgnoreCase;
					// Apply this command against the word
					if ((attr._Token == "~") ||
						(attr._Token.Equals(item.Token, currentComparison)))
					{
						if ((attr._pre == "~") ||
							(item.PreToken.s.EndsWith(attr._pre, currentComparison)))
						{
							if ((attr._preAny == "~") ||
								(item.PreToken.s.IndexOf(attr._preAny, currentComparison) != -1))
							{
								if ((attr._post == "~") ||
									(item.PostToken.s.StartsWith(attr._post, currentComparison)))
								{
									if ((attr._postAny == "~") ||
										(item.PostToken.s.IndexOf(attr._postAny, currentComparison) != -1))
									{
										// Setup to do the Replace
										if (attr._pre != "~")
										{
											startEdit = item.PreToken.s.Length - attr._pre.Length;
											lengthEdit = attr._pre.Length;
										}
										if (attr._preAny != "~")
										{
											startEdit = item.PreToken.s.IndexOf(attr._preAny, currentComparison);
											lengthEdit = attr._preAny.Length;
										}
										if (attr._post != "~")
										{
											startEdit = 0;
											lengthEdit = attr._post.Length;
										}
										if (attr._postAny != "~")
										{
											startEdit = item.PostToken.s.IndexOf(attr._postAny, currentComparison);
											lengthEdit = attr._postAny.Length;
										}

										// Now do the replaces
										if (attr._ReplaceToken != "~")
										{
											item.Token = attr._ReplaceToken;
										}
										if (attr._ReplacePre != "~")
										{
											item.PreToken.s = item.PreToken.s.Remove(startEdit, lengthEdit);
											item.PreToken.s = item.PreToken.s.Insert(startEdit, attr._ReplacePre);
										}
										if (attr._ReplacePost != "~")
										{
											item.PostToken.s = item.PostToken.s.Remove(startEdit, lengthEdit);
											item.PostToken.s = item.PostToken.s.Insert(startEdit, attr._ReplacePost);
										}
									}
								}
							}
						}
					}
				}
			}
		}

		private string GetOutputTextToSpeak ()
		{
			try
			{

				StringBuilder outt = new StringBuilder(4048);
				if (tokens.First.Value.PreToken.s != "")
				{
					outt.Append(tokens.First.Value.PreToken);
				}

				foreach (wordToken item in tokens)
				{
					outt.Append(item.Token);
					outt.Append(item.PostToken);
				}

				return outt.ToString();

			}
			catch (Exception)
			{
				throw;
			}
		}

		/// <summary>Process_MultiChars</summary>
		/// <param name="InTextToSpeak">Input string to process</param>
		private void Process_MultiChars (string regEx, string replace)
		{
			try
			{
				Regex r = new Regex(regEx);
				foreach (wordToken item in tokens)
				{

					// Process PreToken
					string PreItem = item.PreToken.s;
					MatchCollection mc = r.Matches(PreItem);

					if (mc.Count > 0)
					{
						GroupCollection gc = mc[0].Groups;
						if (gc.Count > 0)
						{
							CaptureCollection cc = gc[1].Captures;

							if (cc.Count > 0)
							{
								int counter = cc[0].Value.Length;
								string toReplace = cc[0].Value.Substring(0, 1);

								int startEdit = cc[0].Index;
								int lengthEdit = cc[0].Length;

								item.PreToken.s = item.PreToken.s.Remove(startEdit, lengthEdit);
								item.PreToken.s = item.PreToken.s.Insert(startEdit, toReplace);
							}
						}
					}

					// process PostToken
					string PostItem = item.PostToken.s;
					mc = r.Matches(PostItem);

					if (mc.Count > 0)
					{
						GroupCollection gc = mc[0].Groups;
						if (gc.Count > 0)
						{
							CaptureCollection cc = gc[1].Captures;

							if (cc.Count > 0)
							{
								int counter = cc[0].Value.Length;
								string toReplace = cc[0].Value.Substring(0, 1);

								int startEdit = cc[0].Index;
								int lengthEdit = cc[0].Length;

								item.PostToken.s = item.PostToken.s.Remove(startEdit, lengthEdit);
								item.PostToken.s = item.PostToken.s.Insert(startEdit, toReplace);
							}
						}
					}
				}
			}
			catch (Exception)
			{
				throw;
			}
		}

		/// <summary>Process_Email</summary>
		private void Process_Email (string regEx, string replace)
		{
			// Do nothing at this time
		}

		/// <summary>Process_TimeFix</summary>
		private void Process_TimeFix (string regEx, string replace)
		{
			Regex r = new Regex(@"\d{2}$");
			bool foundHours = false;
			bool foundMinutes = false;
			bool foundSeconds = false;
			// Test 10 11:30 pm 12:30:50 AM 
			try
			{
				LinkedListNode<wordToken> itemNode = tokens.First;
				wordToken item = itemNode.Value;

				do
				{
					Match m = r.Match(item.Token);
					if (m.Success)
					{
						if (foundHours)
						{
							if (item.PreToken.s == ":")
							{
								if (foundMinutes)
								{
									// we found Seconds, lets do something
									item.PreToken.s = "";
									item.Token = "";
									foundHours = false;
									foundMinutes = false;
									foundSeconds = false;
								}
								else
								{
									if (item.PostToken.s == ":")
									{
										// we found minutes
										foundMinutes = true;
									}
									else
									{
										foundHours = false;
										foundMinutes = false;
										foundSeconds = false;
									}
								}
							}
							else
							{
								foundHours = false;
								foundMinutes = false;
								foundSeconds = false;
							}
						}
						else
						{
							// we found a 2 digit number, maybe hours
							if (item.PostToken.s == ":")
							{
								foundHours = true;
							}
						}
					}
					else
					{
						foundHours = false;
						foundMinutes = false;
						foundSeconds = false;
					}

					itemNode = itemNode.Next;
					if (itemNode != null)
					{
						item = itemNode.Value;
					}

				} while (itemNode != null);

			}
			catch (Exception)
			{
				throw;
			}
		}


		private void ParseTextToSpeak (string InTextToSpeak)
		{
			StringBuilder preToken = new StringBuilder();
			StringBuilder Token = new StringBuilder();

			bool inWord = false;
			LinkedListNode<wordToken> theToken =
				new LinkedListNode<wordToken>(new wordToken(new PrePost(""), "", new PrePost("")));
			PrePost lastPreToken = new PrePost("");

			foreach (char theChar in InTextToSpeak)
			{
				// If found delemiter char
				if (wordCharDel.IndexOf(theChar) == -1)
				{
					if (inWord)
					{
						theToken = new LinkedListNode<wordToken>(
							new wordToken(lastPreToken,
										  Token.ToString(),
										  new PrePost("SetLater")));
						tokens.AddLast(theToken);

						preToken.Length = 0;
						preToken.Append(theChar);

						Token.Length = 0;
						inWord = false;
					}
					else
					{
						preToken.Append(theChar);
					}
				}
				// found word char
				else
				{
					if (inWord)
					{
						Token.Append(theChar);
					}
					else
					{
						if (preToken.Length > 0)
						{
							lastPreToken = new PrePost(preToken.ToString());
							if (theToken.Value.PostToken.s.Length > 0)
							{
								theToken.Value.PostToken = lastPreToken;
							}
						}
						Token.Append(theChar);
						inWord = true;
					}
				}
			}

			if (inWord)
			{
				lastPreToken = new PrePost(preToken.ToString());
				theToken = new LinkedListNode<wordToken>(
					new wordToken(lastPreToken,
								  Token.ToString(),
								  new PrePost("")));
				tokens.AddLast(theToken);
			}
			else
			{
				if ((preToken.Length > 0) && (theToken.Value.PostToken.s.Length > 0))
				{
					lastPreToken = new PrePost(preToken.ToString());
					theToken.Value.PostToken = lastPreToken;
				}
			}
		}

		/// <summary>Short Summary of Program.ProcessTextToSpeak</summary>
		private void ProcessTextToSpeak ()
		{
			try
			{
				// ------------------------------------------
				// Process Speak Tags
				// ------------------------------------------
				Process_Replace();

				// ------------------------------------------
				// Process Method Tags before replace
				// ------------------------------------------

				foreach (KeyValuePair<String, bool> item in MethodFlag)
				{
					// test text TheHeaderLine
					if (item.Value)
					{
						switch (item.Key)
						{
							case "MultiWordCaps":
								{
									Process_MultiWordCaps(MethodRegEx[item.Key], MethodReplace[item.Key]);
									break;
								}
							case "DateFix":
								Process_DateFix();
								break;

							default:
								break;
						}
					}
				}

				// ------------------------------------------
				// Process Speak Tags
				// ------------------------------------------
				Process_Replace();

				// ------------------------------------------
				// Process Method Tags
				// ------------------------------------------
				foreach (KeyValuePair<String, bool> item in MethodFlag)
				{
					if (item.Value)
					{
						switch (item.Key)
						{
							//case "MultiWordCaps":
							//    {
							//        Process_MultiWordCaps(MethodRegEx[item.Key], MethodReplace[item.Key]);
							//        break;
							//	  }
							case "MultiChars":
								{
									Process_MultiChars(MethodRegEx[item.Key], "");
									break;
								}
							case "Email":
								{
									Process_Email("", "");
									break;
								}
							case "TimeFix":
								{
									Process_TimeFix("", "");
									break;
								}

							//case "egie":
							//    {
							//        Process_egie("", "");
							//        break;
							//    }

							default:
								break;
						}
					}
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Unexpected error found in ProcessTextToSpeak: " + ex.Message);
			}
		}
		/// <summary>
		/// 4/2, and 4/2/2009 thanks 
		/// </summary>
		private void Process_DateFix ()
		{

			for (LinkedListNode<wordToken> item = tokens.First;
				 item != null;
				 item = item.Next)
			{
				LinkedListNode<wordToken> prev = item.Previous;
				if (prev != null)
				{
					int prevNumber = ConvertInt(prev.Value.Token);
					if (prevNumber != -1)
					{
						if (item.Value.PreToken.s == "/")
						{
							int secondNumber = ConvertInt(item.Value.Token);
							if (secondNumber != -1)
							{
								if (item.Next != null && item.Next.Value.PreToken.s == "/")
								{
									int thirdNumber = ConvertInt(item.Next.Value.Token);
									if (thirdNumber != -1)
									{
										DateTime theDate;
										try
										{
											theDate = new DateTime(thirdNumber, prevNumber, secondNumber);

										}
										catch (Exception)
										{
											continue;
										}
										item.Value.Token = theDate.ToString("MMMM dd yyyy");
										item.Value.PreToken.s = " ";
										item.Value.PostToken.s = " ";
										prev.List.Remove(prev);
										item.List.Remove(item.Next);
									}
								}
								else
								{
									DateTime theDate = new DateTime(2009, prevNumber, secondNumber);
									item.Value.Token = theDate.ToString("MMMM dd");
									item.Value.PreToken.s = " ";
									prev.List.Remove(prev);
								}

							}
						}
					}

					//string postChar = ;
					//string itemNumber = ;
					//string postChar2 = item.Next.Value.PreToken.s;
					//string item2Number = item.Next.Value.Token;
					//DateTime theDate = new DateTime(year, MonthCalendar, Day);
					//string dateTalk = theDate.ToString("mmm dd yyyy");

				}
			}

		}

		private int ConvertInt (string stringin)
		{
			int outInt;
			if (int.TryParse(stringin, out outInt))
			{
				return outInt;
			}
			return -1;
		}

		/// <summary>Short Summary of Program.LoadRules</summary>
		private void LoadRules (string RuleFile)
		{
			WordReplace = new Dictionary<string, string>();
			MethodFlag = new Dictionary<string, bool>();
			MethodReplace = new Dictionary<string, string>();
			MethodRegEx = new Dictionary<string, string>();

			SpeakCmds = new List<SpeakAttr>(200);

			CharReplace = new Dictionary<string, string>();

			XmlDocument WRDoc = new XmlDocument();
			try
			{
				WRDoc.Load(Path.ChangeExtension(AppDomain.CurrentDomain.BaseDirectory + RuleFile, ".xml"));
				// Load Options
				XmlElement xmlOptions = WRDoc.DocumentElement["Options"];
				if (xmlOptions != null)
				{
					string theValue = xmlOptions.GetAttribute("Rate");
					if (theValue != string.Empty)
					{
						voice.Rate = Convert.ToInt32(theValue);
					}

					theValue = xmlOptions.GetAttribute("Volume");
					if (theValue != string.Empty)
					{
						voice.Volume = Convert.ToInt32(theValue);
					}
				}

				XmlNodeList xmlSpeak = WRDoc.GetElementsByTagName("Speak");

				foreach (XmlElement item in xmlSpeak)
				{
					XmlAttributeCollection attrs = item.Attributes;
					SpeakAttr theAttrs = new SpeakAttr();

					for (int i = 0; i < attrs.Count; i++)
					{
						switch (attrs[i].Name)
						{
							case "Pre":
								theAttrs._pre = attrs[i].Value;
								break;
							case "PreAny":
								theAttrs._preAny = attrs[i].Value;
								break;
							case "Post":
								theAttrs._post = attrs[i].Value;
								break;
							case "PostAny":
								theAttrs._postAny = attrs[i].Value;
								break;
							case "Token":
								theAttrs._Token = attrs[i].Value;
								break;
							case "ReplacePre":
								theAttrs._ReplacePre = attrs[i].Value;
								break;
							case "ReplacePost":
								theAttrs._ReplacePost = attrs[i].Value;
								break;
							case "ReplaceToken":
								theAttrs._ReplaceToken = attrs[i].Value;
								break;
							case "CaseSensitive":
								theAttrs._caseSensitive = true;
								break;
							default:
								SpeakText("Error in rule file.");
								break;
						}
					}

					SpeakCmds.Add(theAttrs);
				}

				XmlNodeList xmlMethod = WRDoc.GetElementsByTagName("Method");
				foreach (XmlElement item in xmlMethod)
				{
					MethodFlag[item.GetAttribute("Name")] = (item.GetAttribute("Process") == "true" ? true : false);
					if (item.HasAttribute("RegEx"))
					{
						MethodRegEx[item.GetAttribute("Name")] = item.GetAttribute("RegEx");
					}
					if (item.HasAttribute("Replace"))
					{
						MethodReplace[item.GetAttribute("Name")] = item.GetAttribute("Replace");
					}
				}

			}
			catch (FileNotFoundException)
			{
				SpeakText("Could not find or read the rule file.");
			}
			catch (Exception ex)
			{
				throw new Exception("Unexpected error found in Program.LoadRules : " + ex.Message);
			}
		}

		/// <summary>Short Summary of Program.SpeakText</summary>
		/// <param name="inTextToSpeak">A String containing the text to speak</param>
		private bool SpeakText (string inTextToSpeak)
		{
			EventWaitHandle waitHandleStop = new EventWaitHandle(false, EventResetMode.AutoReset, "GrayT_SayIt_Stop");
			EventWaitHandle waitHandlePauseResume = new EventWaitHandle(false, EventResetMode.AutoReset, "GrayT_SayIt_PauseResume");
			bool pauseIt = false;
			bool currentlyPaused = false;
			bool stopIt = false;

			try
			{
				//voice.Priority = SpeechLib.SpeechVoicePriority.SVPAlert;
				int SpeakId = voice.Speak(inTextToSpeak, speakFlag);
				do
				{
					pauseIt = waitHandlePauseResume.WaitOne(500, false);
					stopIt = waitHandleStop.WaitOne(500, false);

					if (stopIt)
					{
						voice.Speak(", Stopped.", speakFlag);
					}

					if (pauseIt)
					{
						if (!currentlyPaused)
						{
							voice.Pause();
							currentlyPaused = true;
							continue;
						}
						else
						{
							voice.Resume();
							currentlyPaused = false;
							continue;
						}
					}

				}
				while (voice.Status.RunningState != SpeechLib.SpeechRunState.SRSEDone);

				//voice.WaitUntilDone(Timeout.Infinite);

				return true;
			}
			catch (Exception)
			{
				return false;
			}
			finally
			{
				waitHandleStop.Close();
				waitHandlePauseResume.Close();
			}
		}

		/// <summary>Short Summary of Program.WriteText</summary>
		/// <param name="inTextToSpeak">A String containing the text to speak</param>
		private bool WriteText (string inTextToSpeak)
		{
			try
			{
				SpeechLib.SpFileStream speakStream = new SpeechLib.SpFileStream();
				speakStream.Format.Type = SpeechLib.SpeechAudioFormatType.SAFT22kHz16BitMono;

				speakStream.Open(AppDomain.CurrentDomain.BaseDirectory + @"\" + outputFilename, SpeechLib.SpeechStreamFileMode.SSFMCreateForWrite, false);
				voice.AllowAudioOutputFormatChangesOnNextSet = false;
				voice.AudioOutputStream = speakStream;

				SpeechLib.ISpeechObjectTokens tokens = voice.GetVoices("", "");
				voice.Voice = tokens.Item(0);

				voice.Rate = 8;
				//voice.EventInterests = SpeechLib.SpeechVoiceEvents.SVESentenceBoundary;
				//voice.Sentence += new SpeechLib._ISpeechVoiceEvents_SentenceEventHandler(sentence_Event);

				speakFlag = SpeechLib.SpeechVoiceSpeakFlags.SVSFDefault
					| SpeechLib.SpeechVoiceSpeakFlags.SVSFlagsAsync
					| SpeechLib.SpeechVoiceSpeakFlags.SVSFPurgeBeforeSpeak
					//| SpeechLib.SpeechVoiceSpeakFlags.SVSFNLPSpeakPunc
					//| SpeechLib.SpeechVoiceSpeakFlags.SVSFIsXML;
					;

				voice.Speak(inTextToSpeak, speakFlag);
				voice.WaitUntilDone(Timeout.Infinite);

				speakStream.Close();
				speakStream = null;

				SpeechLib.ISpeechObjectTokens Atokens = voice.GetAudioOutputs(null, "");
				voice.AudioOutput = Atokens.Item(0);
				voice.AudioOutputStream.Format.Type = SpeechLib.SpeechAudioFormatType.SAFT22kHz16BitMono;
				voice.AudioOutputStream = voice.AudioOutputStream;

				voice.Speak("Speech has been completly written to the file, " + outputFilename, speakFlag);
				voice.WaitUntilDone(Timeout.Infinite);

				return true;
			}
			catch (Exception)
			{
				return false;
			}
		}

		/// <summary>InitOptions: Method to process the SayIt command line options</summary>
		/// <param name="args">CommandOptions: Options to process</param>
		private void InitOptions (CommandOptions args)
		{
			try
			{
				String Command = string.Empty;
				String[] tokens = null;

				args.getArg(out Command, out tokens);
				while (Command != "end")
				{
					switch (Command)
					{
						case "data":
							RuleName = tokens[0];
							break;
						case "-v":
							voice.Volume = Convert.ToInt32(tokens[0]);
							break;
						case "-r":
							voice.Rate = Convert.ToInt32(tokens[0]);
							break;
						case "-f":
							writeToFile = true;
							outputFilename = tokens[0];
							break;
						case "-set":
							break;
						default:
							break;
					}
					args.getArg(out Command, out tokens);

				}
			}
			catch (Exception)
			{
				throw;
			}
		}

		/// <summary>CommandOptions: Class to parse the command line options</summary>
		class CommandOptions
		{
			private String[] _args;
			private int CurrPointer = 0;

			public CommandOptions (string[] inArgs)
			{
				_args = inArgs;
			}

			public void getArg (out string Command, out string[] Params)
			{
				Params = null;
				if (_args.Length == 0 || CurrPointer + 1 > _args.Length)
				{
					Command = "end";
					Params = null;
					return;
				}
				else
				{
					string token = _args[CurrPointer];
					if (token.Substring(0, 1) == "-")
					{
						Command = token;
						CurrPointer += 1;
						List<String> tokens = new List<string>();

						for (int i = CurrPointer; i < _args.Length; i++)
						{
							if (_args[i].Substring(0, 1) != "-")
							{
								tokens.Add(_args[i]);
								CurrPointer += 1;
							}
							else
							{
								Params = new string[tokens.Count];
								tokens.CopyTo(Params);
								return;
							}
						}
						Params = new string[tokens.Count];
						tokens.CopyTo(Params);
						//CurrPointer += 1;
						return;
					}
					else
					{
						Command = "data";
						Params = new string[1] { token };
						CurrPointer += 1;
						return;
					}
				}
			}
		}

		/// <summary>Short Summary of Program.Main</summary>
		/// <param name="args">Command Line Arguments</param>
		[STAThread]
		static void Main (string[] args)
		{
			//Application.EnableVisualStyles();
			//Application.SetCompatibleTextRenderingDefault(false);
			//Application.Run(new FormSayIt());

			//args = new string[] { "TonyGray", "-f", "output.wav" };

			SpeakItClass speak = new SpeakItClass();
			string textToSpeak = string.Empty;
			string processedTextToSpeak = string.Empty;

			if (speak.InitSpeech())
			{
				CommandOptions options = new CommandOptions(args);
				speak.InitOptions(options);
				if (speak.RuleName.ToLower() == "stop")
				{
					EventWaitHandle waitHandleStop = new EventWaitHandle(false, EventResetMode.AutoReset, "GrayT_SayIt_Stop");
					waitHandleStop.Set();
					do
					{
						Thread.Sleep(3000);
					} while (waitHandleStop.WaitOne(500, false));
					waitHandleStop.Close();
				}
				else if (speak.RuleName.ToLower() == "pauseresume")
				{
					EventWaitHandle waitHandlePauseResume = new EventWaitHandle(false, EventResetMode.AutoReset, "GrayT_SayIt_PauseResume");
					waitHandlePauseResume.Set();
					do
					{
						Thread.Sleep(3000);
					} while (waitHandlePauseResume.WaitOne(500, false));
					waitHandlePauseResume.Close();
				}
				else
				{
					speak.LoadRules(speak.RuleName);
					if (speak.GetTextToSpeak(1, out textToSpeak))
					{
						//textToSpeak = @"..Word1.\Word2./";
						speak.ParseTextToSpeak(textToSpeak);
						speak.ProcessTextToSpeak();
						processedTextToSpeak = speak.GetOutputTextToSpeak();

						if (speak.writeToFile)
						{
							if (!speak.WriteText(processedTextToSpeak))
							{
								speak.SpeakText("Error writting spoken text to file!");
							}
						}
						else
						{
							if (!speak.SpeakText(processedTextToSpeak))
							{
								speak.SpeakText("Error speaking Text.");
							}
						}
					}
					else
					{
						textToSpeak = "I didn't find anything to say.";
					}
				}
			}
		}
	}
}
