private async Task<Boolean> RegisterOrRenewSubscription(Resource resource)
{
	OutlookGraph.LoginInformation loginInformation = 
		OutlookHelper.getLoginInformations(resource);
	OutlookGraph? graph = 
		await OutlookHelper.getOrCreateGraphInstance(loginInformation, Tenant.Uid);
	Subscription? requestBody = new Subscription
	{
		ChangeType = "created,updated",
		NotificationUrl = Tenant.FullUrlExtern+"/api/outlook/notify",
		LifecycleNotificationUrl = Tenant.FullUrlExtern+"/api/outlook/lifecycle_notify",
		Resource = "/users/" + graph.userHandle + "/events",
		ExpirationDateTime = DateTimeOffset.Now.AddDays(2),
		ClientState = "SecretClientState",
	};
	SubscriptionCollectionResponse? subscriptions = 
		await graph.graphClient.Subscriptions.GetAsync();
	if (subscriptions != null)
	{
		// Delete all old subscriptions for this resource
		for (int i = 0; i < subscriptions.Value.Count; i++)
		{
			if (subscriptions.Value[i] == null || subscriptions.Value[i].Id == null)
			{
				continue;
			}
			await graph.graphClient.Subscriptions[subscriptions.Value[i].Id].DeleteAsync();
		}
	}
	// Create the new subscription
	Subscription? result = await graph.graphClient.Subscriptions.PostAsync(requestBody);
	if (result != null)
	{
		if (resource.DataJson.TryGetValue("outlook_subscription_id", out _))
		{
			resource.DataJson.Remove("outlook_subscription_id");
		}
		// Write the "outlook_subscription_id" in the resource to keep it persistent
		resource.DataJson.TryAdd("outlook_subscription_id", result.Id);
		_resourceManager.Save(resource);
		return true;
	}
	else
	{
		return false;
	}
}