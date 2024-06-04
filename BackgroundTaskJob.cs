protected async Task<Boolean> executeInternalAsync(String? body)
{
	if (body.IsNullOrWhiteSpace())
	{
		// No body passed
		return false;
	}
	JObject parsedBody;
	try
	{
		parsedBody = JObject.Parse(body);
	}
	catch
	{
		// Could not parse body
		return false;
	}
	String? outlook_event_id = parsedBody["value"][0]["resourceData"]["id"].ToString();
	String? subscription_id = parsedBody["value"][0]["subscriptionId"].ToString();
	String? changeType = parsedBody["value"][0]["changeType"].ToString();
	if (outlook_event_id.IsNullOrWhiteSpace() || subscription_id.IsNullOrWhiteSpace() || changeType.IsNullOrWhiteSpace())
	{
		// Parsed body does not contain all necessary values
		return false;
	}
	Resource? resource = _resourceManager.GetResourcesByOutlookSubscriptionId(subscription_id).FirstOrDefault();
	if (resource == null)
	{
		// No resource is referenced with the provided subscription
		return false;
	}
	OutlookGraph.LoginInformation loginInformation = OutlookHelper.getLoginInformations(resource);
	OutlookGraph? graph = await OutlookHelper.getOrCreateGraphInstance(loginInformation, Tenant.Uid);
	OutlookEvent? outlookEvent = await graph.getEvent(outlook_event_id);
	if (outlookEvent == null)
	{
		// Graph categorizes deletions as changes.
		// To differentiate, check if the event is null.
		// It is null if not available online
		changeType = "deleted";
	}
	switch (changeType)
	{
		case "created":
			// Wait until nothing blocks the OutlookHelper
			if (Utilities.OutlookHelper.accessSynchronizationTasks == null)
			{
				Utilities.OutlookHelper.accessSynchronizationTasks = new SemaphoreSlim(1, 1);
			}
			await Utilities.OutlookHelper.accessSynchronizationTasks.WaitAsync();
			try
			{
				List<Appointment> possibleCorrespondingAppointments = _appointmentManager
					.GetAppointmentsByOutlookEventId(Tenant, resource.Uid, outlook_event_id);
				if (possibleCorrespondingAppointments == null || possibleCorrespondingAppointments.Count() == 0)
				{
					// Create an appointment with "Event" and placeholder data
					Appointment placeholderAppointment = new Appointment(Tenant.Uid);
					placeholderAppointment.Resource = resource;
					placeholderAppointment.AppointmentDatetime = outlookEvent.start;
					placeholderAppointment.AppointmentEndDatetime = outlookEvent.end;
					placeholderAppointment.AppointmentDuration = (int)(outlookEvent.end - outlookEvent.start).TotalMinutes;
					JObject? detailsJson = Newtonsoft
						.Json
						.JsonConvert
						.DeserializeObject<JObject>(
							"{\"booked_by\": \"outlook\", \"comment\": \"Outlook\",\"outlook_event_id\": \"" + outlook_event_id + "\"}"
							);
					if ( detailsJson == null )detailsJson = new JObject();
					placeholderAppointment.DetailsJson = detailsJson;
					placeholderAppointment.AppointmentNumber = "Outlook";
					placeholderAppointment.AppointmentStatus = AppointmentStatus.reserved;
					_appointmentManager.Save(placeholderAppointment);
					_resourceAvailabilityManager.UpdateAvailabilityByAppointment(placeholderAppointment);
				}
			}
			finally
			{
				Utilities.OutlookHelper.accessSynchronizationTasks.Release();
			}
			break;
		case "updated":
			// Find the corresponding appointment or blocker and change the timeframe
			List<Appointment> appointments_to_update = _appointmentManager
				.GetAppointmentsByOutlookEventId(Tenant, resource.Uid, outlook_event_id);
			if (appointments_to_update != null)
			{
				for (int i = 0; i < appointments_to_update.Count(); i++)
				{
					Appointment appointment = appointments_to_update[i];
					appointment.AppointmentDatetime = outlookEvent.start;
					appointment.AppointmentEndDatetime = outlookEvent.end;
					appointment.AppointmentDuration = (int)(outlookEvent.end - outlookEvent.start).TotalMinutes;
					appointment.DetailsJson.TryAdd("outlook_event_id", outlook_event_id);
					appointment.AppointmentNumber = "Outlook";
					_appointmentManager.Save(appointment);
					_resourceAvailabilityManager.UpdateAvailabilityByAppointment(appointment);
				}
			}
			else
			{
				// No appointment is associated with the notification
				return false;
			}
			break;
		case "deleted":
			List<Appointment> appointments_to_delete = _appointmentManager
				.GetAppointmentsByOutlookEventId(Tenant, resource.Uid, outlook_event_id);
			if (appointments_to_delete != null)
			{
				for (int i = 0; i < appointments_to_delete.Count(); i++)
				{
					Appointment appointment = appointments_to_delete[i];
					_appointmentManager.CancelAppointment(appointment, "Outlook synchronisation", true);
				}
			}
			else
			{
				// No appointment is associated with the notification
				return false;
			}
			break;
		default:
			// invalid notification changeType
			return false;
	}
	return true;
}