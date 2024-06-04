using Rocket.Models;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json.Linq;
using Rocket.Manager.Interfaces;
using Rocket.Manager;
using System.Collections.Concurrent;

namespace Rocket.Utilities
{
	public class OutlookHelper
	{
		public static ConcurrentDictionary<Guid, List<OutlookGraph>> connected = new ConcurrentDictionary<Guid, List<OutlookGraph>>();
		public static SemaphoreSlim accessSynchronizationTasks = new SemaphoreSlim(1, 1);
		public class OutlookGraph
		{
			public GraphServiceClient graphClient;
			public String userHandle;
			public OutlookGraph(LoginInformation loginInformation)
			{
				this.userHandle = loginInformation.userHandle;
				var scopes = new[] { "https://graph.microsoft.com/.default" };
				var options = new ClientSecretCredentialOptions
				{
					AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
				};
				var clientSecretCredential = new ClientSecretCredential(
						loginInformation.tenantId,
						loginInformation.clientId,
						loginInformation.clientSecret,
						options
						);
				graphClient = new GraphServiceClient(clientSecretCredential, scopes);
			}
			public async Task<Boolean> checkToken()
			{
				try
				{
					// Simple always succeeding call to test the authorization
					await graphClient.Users[userHandle].GetAsync();
				}
				catch
				{
					return false;
				}
				return true;
			}
			private async Task<OutlookEvent?> readEvent(String uid)
			{
				try
				{
					var result = await graphClient.Users[userHandle].Calendar.Events[uid].GetAsync((requestConfiguration) =>
					{
						requestConfiguration.QueryParameters.Select = new String[] { "subject", "start", "end", "body", "location", "showas" };
						requestConfiguration.Headers.Add("Prefer", "outlook.timezone=\"Europe/Berlin\"");
					});
					String body = result?.Body?.Content ?? "";
					return new OutlookEvent(result?.Id ?? "", result?.Subject ?? "", body, result?.Start?.ToDateTime() ?? new DateTime(), result?.End?.ToDateTime() ?? new DateTime());
				}
				catch
				{
					return null;
				}
			}
			private async Task<List<OutlookEvent>> readEvents(DateTime date)
			{
				String sDate = date.Year + "-" + date.Month + "-" + date.Day;
				var result = await graphClient.Users[userHandle].Calendar.Events.GetAsync((requestConfiguration) =>
				{
					requestConfiguration.QueryParameters.Select = new String[] { "subject", "start", "end", "location", "body" };
					requestConfiguration.QueryParameters.Filter = " start/dateTime ge '" + sDate + "T00:00:00Z' AND start/dateTime le '" + sDate + "T23:59:59Z'";
					requestConfiguration.Headers.Add("Prefer", "outlook.timezone=\"Europe/Berlin\"");//Change to tenant timezone in the future
				});
				List<OutlookEvent> outlookEvents = new List<OutlookEvent>();
				if (result != null && result.Value != null)
				{
					for (int i = 0; i < result.Value.Count; i++)
					{
						if (result.Value[i] == null || result.Value[i].Id == null || result.Value[i].Subject == null)
						{
							continue;
						}
						// Convert the Events to OutlookEvents and replace missing values with placeholders
						outlookEvents.Add(new OutlookEvent(result?.Value[i]?.Id ?? "", result?.Value[i]?.Subject ?? "", result?.Value[i]?.Body?.Content ?? "", result?.Value[i]?.Start?.ToDateTime() ?? new DateTime(), result?.Value[i]?.End?.ToDateTime() ?? new DateTime()));
					}
				}
				return outlookEvents;
			}
			public async Task<OutlookEvent?> getEvent(String uid)
			{
				// check token
				if (!await checkToken())
				{
					throw new Exception("Unauthorized");
				}
				OutlookEvent? outlookEvent = await readEvent(uid);
				return outlookEvent;
			}
			public async Task<List<OutlookEvent>> getEvents(DateTime date)
			{
				// check token
				if (!await checkToken())
				{
					throw new Exception("Unauthorized");
				}
				List<OutlookEvent> outlookEvents = await readEvents(date);
				return outlookEvents;
			}
			private async Task<Boolean> isSlotFree(DateTime start, DateTime end)
			{
				List<OutlookEvent> outlookEvents = await readEvents(start);
				for (int i = 0; i < outlookEvents.Count; i++)
				{
					if (outlookEvents[i].start >= start && outlookEvents[i].start < end || outlookEvents[i].end > start && outlookEvents[i].end <= end)
					{
						return false;
					}
				}
				return true;
			}
			public async Task<String> createEvent(OutlookEvent outlookEvent, Boolean allowConcurrentOutlookAppointments = true)
			{
				if (!await checkToken())
					throw new Exception("Unauthorized");
				if (!outlookEvent.IsValid()) return "";
				if (!await isSlotFree(outlookEvent.start, outlookEvent.end) && !allowConcurrentOutlookAppointments) { return ""; }
				var requestBody = new Event
				{
					Subject = outlookEvent.subject,
					Body = new ItemBody
					{
						ContentType = BodyType.Html,
						Content = outlookEvent.body,
					},
					Start = new DateTimeTimeZone
					{
						DateTime = outlookEvent.start.ToString("yyyy-MM-ddTHH:mm:ss"),
						TimeZone = "Europe/Berlin",
					},
					End = new DateTimeTimeZone
					{
						DateTime = outlookEvent.end.ToString("yyyy-MM-ddTHH:mm:ss"),
						TimeZone = "Europe/Berlin",
					},
					AllowNewTimeProposals = false,
				};
				Event? result = await graphClient.Users[userHandle].Calendar.Events.PostAsync(requestBody, (requestConfiguration) =>
				{
					requestConfiguration.Headers.Add("Prefer", "outlook.timezone=\"Europe/Berlin\"");
				});
				if (result != null && result.Id != null)
				{
					return result.Id;
				}
				return "";
			}
			public async Task<Boolean> updateEvent(String uid, OutlookEvent outlookEvent)
			{
				if (!await checkToken()) return false;
				if (!outlookEvent.IsValid()) return false;
				if (!await isSlotFree(outlookEvent.start, outlookEvent.end)) { return false; }
				var requestBody = new Event
				{
					Subject = outlookEvent.subject,
					Body = new ItemBody
					{
						ContentType = BodyType.Html,
						Content = outlookEvent.body,
					},
					Start = new DateTimeTimeZone
					{
						DateTime = outlookEvent.start.ToString("yyyy-MM-ddTHH:mm:ss"),
						TimeZone = "Europe/Berlin",
					},
					End = new DateTimeTimeZone
					{
						DateTime = outlookEvent.end.ToString("yyyy-MM-ddTHH:mm:ss"),
						TimeZone = "Europe/Berlin",
					},
					AllowNewTimeProposals = false,
				};
				var result = await graphClient.Users[userHandle].Calendar.Events[uid].PatchAsync(requestBody, (requestConfiguration) =>
				{
					requestConfiguration.Headers.Add("Prefer", "outlook.timezone=\"Europe/Berlin\"");
				});
				return true;
			}
			public async Task<Boolean> deleteEvent(String uid)
			{
				if (!await checkToken()) return false;
				try
				{
					await graphClient.Users[userHandle].Calendar.Events[uid].DeleteAsync();
				}
				catch
				{
					return false;
				}
				return true;
			}
			public class LoginInformation
			{
				public String tenantId { get; set; }
				public String clientId { get; set; }
				public String clientSecret { get; set; }
				public String userHandle { get; set; }
				public LoginInformation(String tenantId, String clientId, String clientSecret, String userHandle)
				{
					this.tenantId = tenantId;
					this.clientId = clientId;
					this.clientSecret = clientSecret;
					this.userHandle = userHandle;
				}
			}
		}
		public class OutlookEvent
		{
			public String id;
			public String subject;
			public String body;
			public DateTime start;
			public DateTime end;
			public OutlookEvent(String id, String subject, String body, DateTime start, DateTime end)
			{
				this.id = id;
				this.subject = subject;
				this.body = body;
				this.start = start;
				this.end = end;
			}
			public OutlookEvent()
			{
				this.id = "";
				this.subject = "";
				this.body = "";
				this.start = new DateTime();
				this.end = new DateTime();
			}
			public Boolean IsValid()
			{
				if (this.id == "") return false;
				if (this.subject == "") return false;
				if (this.start == this.end) return false;
				return true;
			}
		}
		static public OutlookGraph.LoginInformation getLoginInformations(Appointment appointment)
		{
			if (appointment.Resource == null || appointment.Tenant == null || appointment.Tenant.ApplicationSettings == null)
			{
				return null;
			}
			return getLoginInformations(appointment.Resource);
		}
		static public OutlookGraph.LoginInformation getLoginInformations(Resource resource)
		{
			if (resource == null || resource.Tenant == null || resource.Tenant.ApplicationSettings == null)
			{
				return null;
			}
			String? tenantID = resource.Tenant.ApplicationSettings.First(setting => setting.SettingName == "rc_module_appointment__AzureTenantID").SettingValue;
			String? clientId = resource.Tenant.ApplicationSettings.First(setting => setting.SettingName == "rc_module_appointment__AzureClientID").SettingValue;
			String? clientSecret = resource.Tenant.ApplicationSettings.First(setting => setting.SettingName == "rc_module_appointment__AzureClientSecret").SettingValue;
			String? userHandle = null;
			if (resource.DataJson.TryGetValue("custom", out JToken? custom) && custom.ToObject<JObject>().TryGetValue("userHandle", out JToken? TuserHandle))
			{
				userHandle = TuserHandle.ToString();
			}
			if (tenantID == null || clientId == null || clientSecret == null || userHandle == null)
			{
				return null;
			}
			return new OutlookGraph.LoginInformation(tenantID, clientId, clientSecret, userHandle);
		}
		public async static Task<OutlookGraph> getOrCreateGraphInstance(OutlookGraph.LoginInformation loginInformation, Guid tenantUid)
		{
			connected.TryGetValue(tenantUid, out List<OutlookGraph>? graphs);
			if (graphs == null)
			{
				graphs = new List<OutlookGraph>();
				connected.TryAdd(tenantUid, graphs);
			}
			OutlookGraph? graph = graphs.Find(client => client.userHandle == loginInformation.userHandle);
			if (graph == null)
			{
				graph = new OutlookGraph(loginInformation);
				graphs.Add(graph);
			}
			if (!await graph.checkToken())
			{
				graphs.Remove(graph);
				graph = new OutlookGraph(loginInformation);
				graphs.Add(graph);
			}
			return graph;
		}
		public static async Task<Appointment> deleteAppointmentInOutlook(Appointment appointment)
		{
			if (appointment.DetailsJson.TryGetValue("outlook_event_id", out JToken? outlook_id))
			{
				OutlookGraph.LoginInformation loginInformation = getLoginInformations(appointment);
				if (loginInformation != null)
				{
					OutlookGraph instantiatedGraph = await getOrCreateGraphInstance(loginInformation, appointment.TenantUid);
					await instantiatedGraph.deleteEvent(outlook_id.ToString());
					appointment.DetailsJson.Remove("outlook_event_id");
				}
			}
			return appointment;
		}
		private static String? getSettingsEntry(Template? template, String entry)
		{
			if (template != null && template.DataJson.TryGetValue("settings", out JToken? Tsettings))
			{
				JObject? settings = Tsettings.ToObject<JObject>();
				if (settings != null && settings.TryGetValue(entry, out JToken? token))
				{
					String? entryContent = token.ToString();
					if (entryContent != null && entryContent != "")
					{
						return entryContent;
					}
				}
			}
			return null;
		}
		private static String getTemplateEntry(TemplateManager templateManager, Guid guid, String entryName, String language)
		{
			String index = entryName;
			String localized_index = index + "_" + language;
			Template? template = templateManager.GetTemplate(guid);
			String? content = null;
			content = getSettingsEntry(template, localized_index);
			if (content != null)
			{
				return content;
			}
			content = getSettingsEntry(template, index);
			if (content != null)
			{
				return content;
			}
			return "";
		}
		public static async Task<Appointment> sendAppointmentToOutlook(Appointment appointment, ITranslationManager translationManager, LiquidRenderContext rctx, Rocket.Data.ApplicationDbContext dbContext)
		{
			rctx.AddObject("appointment_number", appointment.AppointmentNumber);
			String eventContent = "";
			String eventSubject = "";
			Boolean allowConcurrentOutlookAppointments = false;
			if (appointment.Tenant.ApplicationSettings != null)
			{
				// Get the setting whether or not to allow concurrent events in Outlook 
				String? SallowConcurrentOutlookAppointments = appointment.Resource.Tenant.ApplicationSettings.First(setting => setting.SettingName == "rc_module_appointment__AllowConcurrentOutlookAppointments").SettingValue;
				allowConcurrentOutlookAppointments = SallowConcurrentOutlookAppointments != null && SallowConcurrentOutlookAppointments == "true";
				String? templateUID = null;
				// Get template UID from resource
				if (appointment.Resource != null && appointment.Resource.DataJson.TryGetValue("custom", out JToken? custom) && custom.ToObject<JObject>().TryGetValue("templateUID", out JToken? TtemplateUid))
				{
					templateUID = TtemplateUid.ToString();
				}
				if (templateUID != null)
				{
					String dummy = rctx.TemplateMarkup;
					TemplateManager templateManager = new TemplateManager(dbContext);
					rctx.TemplateMarkup = getTemplateEntry(templateManager, new Guid(templateUID), "mail_body", appointment.Tenant.LanguageDefault);
					eventContent = rctx.Render(appointment.Tenant.LanguageDefault);
					rctx.TemplateMarkup = getTemplateEntry(templateManager, new Guid(templateUID), "mail_subject", appointment.Tenant.LanguageDefault);
					eventSubject = rctx.Render(appointment.Tenant.LanguageDefault);
					rctx.TemplateMarkup = dummy;
				}
			}
			OutlookEvent outlookEvent = new OutlookEvent(
					"NEW",
					eventSubject,
					eventContent,
					appointment.AppointmentDatetime.DateTime,
					appointment.AppointmentEndDatetime.HasValue ? appointment.AppointmentEndDatetime.Value.DateTime : appointment.AppointmentDatetime.DateTime
			);
			OutlookGraph.LoginInformation loginInformation = getLoginInformations(appointment);
			if (loginInformation != null)
			{
				OutlookGraph instantiatedGraph = await getOrCreateGraphInstance(loginInformation, appointment.TenantUid);
				String id = await instantiatedGraph.createEvent(outlookEvent,allowConcurrentOutlookAppointments);
				if (id != "")
				{
					appointment.DetailsJson.TryAdd("outlook_event_id", id);
				}
			}
			return appointment;
		}
	}
}
