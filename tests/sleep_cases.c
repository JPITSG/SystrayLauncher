static void BeginMainNavigationTitle(HWND hwnd, UINT64 navigationId) {
    g_mainNavigationLoading=TRUE;
}
static void KickMainWebViewComposition(void) {
    ++kicks;
    ResumeMainWebViewRuntime();
    g_controllerVisible=-1;
    SetMainWebViewControllerVisible(TRUE);
}
static void finish_suspend(size_t index, HRESULT error, BOOL result) {
    assert(index<pendingCount && pending[index]);
    ICoreWebView2TrySuspendCompletedHandler *handler=pending[index]; pending[index]=NULL;
    handler->lpVtbl->Invoke(handler,error,result); handler->lpVtbl->Release(handler);
}
static void expire_timer(UINT id) {
    ticks+=timers[id];
    fire_timer(id);
}
static void init(void) {
    windows[0]=(Window){.visible=TRUE,.pid=1,.rect={0,0,20,20},.frame={0,0,20,20}};
    windows[1]=(Window){.visible=TRUE,.pid=2,.rect={0,0,20,20},.frame={0,0,20,20}};
    zorder[0]=&windows[1]; zorder[1]=g_hwnd; windowCount=2;
    g_webView=&web; g_webViewController=&controller;
    g_isInitialized=TRUE; g_initialPreloadComplete=TRUE; g_mainNavigationLoading=FALSE;
    g_sleepWhenInactive=TRUE; g_webViewDesiredActive=TRUE; g_webViewDesiredVisible=TRUE;
    g_controllerVisible=TRUE; g_jsVisibility=JS_VISIBILITY_SHOWN;
    StartVisibilityTracking(g_hwnd);
}
static void event(DWORD type, HWND hwnd, LONG objectId) {
    VisibilityWinEventProc(NULL,type,hwnd,objectId,CHILDID_SELF,0,0);
}
static void coverage(void) {
    init();
    windows[1].frame.right=10;
    UpdateJsVisibilityState(g_hwnd);
    assert(rendered && !suspends);
    windows[1].frame.right=20;
    UpdateJsVisibilityState(g_hwnd);
    assert(!rendered && suspends==1 && g_webViewSuspendPending);
    finish_suspend(0,S_OK,TRUE);
    assert(g_webViewSuspended);
    unsigned before=scans;
    windows[1].frame.right=10;
    event(EVENT_OBJECT_LOCATIONCHANGE,&windows[1],OBJID_WINDOW);
    assert(scans==before && timers[ID_TIMER_VISIBILITY_CHECK]==50);
    fire_timer(ID_TIMER_VISIBILITY_CHECK);
    assert(rendered && !g_webViewSuspended && !g_webViewSuspendPending);
    assert(g_jsVisibility==JS_VISIBILITY_SHOWN && !timers[ID_TIMER_VISIBILITY_CHECK]);

    windows[2]=windows[1]; windows[2].pid=3; windows[2].frame=(RECT){10,0,20,20};
    zorder[1]=&windows[2]; zorder[2]=g_hwnd; windowCount=3;
    assert(!IsWindowActuallyVisible(g_hwnd)); // Union of two opaque windows.
    windows[1].frame.right=9;
    assert(IsWindowActuallyVisible(g_hwnd)); // A one-pixel exposed strip.
    windows[1].frame.right=20;
    foreground=g_hwnd; before=scans;
    assert(IsWindowActuallyVisible(g_hwnd) && scans==before);
    assert(!liveRegions);
}
static void geometry(void) {
    init();
    for (int style=1;style<=4;style*=2) {
        windows[1].style=style; assert(IsWindowActuallyVisible(g_hwnd));
    }
    windows[1].style=0; windows[1].shaped=TRUE;
    assert(IsWindowActuallyVisible(g_hwnd));
    windows[1].shaped=FALSE; windows[1].cloaked=TRUE;
    assert(IsWindowActuallyVisible(g_hwnd));
    windows[1].cloaked=FALSE; windows[1].minimized=TRUE;
    assert(IsWindowActuallyVisible(g_hwnd));
    windows[1].minimized=FALSE; windows[1].dwmFails=TRUE;
    assert(IsWindowActuallyVisible(g_hwnd));
    windows[1].dwmFails=FALSE; windows[1].frame.right=19; // Invisible border is not coverage.
    assert(IsWindowActuallyVisible(g_hwnd));
    windows[1].frame.right=20;
    regionFails=TRUE; assert(IsWindowActuallyVisible(g_hwnd)); regionFails=FALSE;
    combineFails=TRUE; assert(IsWindowActuallyVisible(g_hwnd)); combineFails=FALSE;
    enumerationFails=TRUE; assert(IsWindowActuallyVisible(g_hwnd)); enumerationFails=FALSE;
    monitorMissing=TRUE; assert(IsWindowActuallyVisible(g_hwnd)); monitorMissing=FALSE;
    windowCount=1; assert(IsWindowActuallyVisible(g_hwnd)); // Target missing from enumeration.
    windowCount=2; assert(!IsWindowActuallyVisible(g_hwnd));
    windows[0].dwmFails=TRUE; assert(IsWindowActuallyVisible(g_hwnd));
    assert(!liveRegions);
}
static void events(void) {
    init(); windows[1].frame.right=10;
    UpdateJsVisibilityState(g_hwnd);
    unsigned beforeScans=scans, beforeTimers=timerSets, beforeShows=shows, beforeBounds=boundsSets;
    windows[2].child=TRUE;
    for (int i=0;i<10000;++i) {
        event(EVENT_OBJECT_LOCATIONCHANGE,&windows[1],-8); // Caret.
        event(EVENT_OBJECT_LOCATIONCHANGE,&windows[2],OBJID_WINDOW); // Child control.
        event(EVENT_OBJECT_LOCATIONCHANGE,NULL,OBJID_WINDOW);
    }
    assert(scans==beforeScans && timerSets==beforeTimers);
    for (int i=0;i<1000;++i) event(EVENT_OBJECT_LOCATIONCHANGE,&windows[1],OBJID_WINDOW);
    assert(timerSets==beforeTimers+1 && scans==beforeScans);
    assert(timers[ID_TIMER_VISIBILITY_CHECK]==10000);
    fire_timer(ID_TIMER_VISIBILITY_CHECK);
    assert(!timers[ID_TIMER_VISIBILITY_CHECK]);
    assert(shows==beforeShows && boundsSets==beforeBounds);
    beforeScans=scans;
    fire_timer(ID_TIMER_VISIBILITY_CHECK); // Already-queued stale WM_TIMER.
    assert(scans==beforeScans);
    QueueVisibilityCheck(FALSE); QueueVisibilityCheck(TRUE);
    assert(timers[ID_TIMER_VISIBILITY_CHECK]==50);
    unsigned armed=timerSets;
    for (int i=0;i<1000;++i) QueueVisibilityCheck(TRUE);
    assert(timerSets==armed);
    fire_timer(ID_TIMER_VISIBILITY_CHECK);
    timerFails=TRUE; QueueVisibilityCheck(TRUE);
    assert(!g_visibilityHooksAvailable && lastPost==WM_APP_VISIBILITY_WAKE);
    windows[1].frame.right=20; assert(IsWindowActuallyVisible(g_hwnd));
}
static void lifecycle(void) {
    init(); assert(g_visibilityTracking && hookSets==8);
    StartVisibilityTracking(g_hwnd); assert(hookSets==8);
    IsWindowActuallyVisible(g_hwnd); assert(g_visibilityLocationHookCount==1);
    QueueVisibilityCheck(TRUE);
    windows[0].visible=FALSE; StartVisibilityTracking(g_hwnd);
    assert(!g_visibilityTracking && !timers[ID_TIMER_VISIBILITY_CHECK]);
    assert(!g_visibilityLocationHookCount);
    for (int i=0;i<512;++i) assert(!hooks[i].active);
    assert(!IsWindowActuallyVisible(g_hwnd));
    windows[0].visible=TRUE; windows[0].minimized=TRUE;
    StartVisibilityTracking(g_hwnd); assert(!g_visibilityTracking);
    windows[0].minimized=FALSE; g_sleepWhenInactive=FALSE;
    StartVisibilityTracking(g_hwnd); assert(!g_visibilityTracking);
    wcscpy(g_config.onHideJs,L"hidden()");
    StartVisibilityTracking(g_hwnd); assert(g_visibilityTracking);
    StopVisibilityTracking(g_hwnd); failHookAt=(int)hookSets+2;
    StartVisibilityTracking(g_hwnd);
    assert(!g_visibilityTracking && IsWindowActuallyVisible(g_hwnd));
    failHookAt=0; StartVisibilityTracking(g_hwnd);
    failHookAt=(int)hookSets+1; // Location hook fails: cannot safely sleep under this window.
    assert(IsWindowActuallyVisible(g_hwnd));
    failHookAt=0;
    for (int i=1;i<70;++i) {
        windows[i]=windows[1]; windows[i].pid=(DWORD)i+1; zorder[i-1]=&windows[i];
    }
    zorder[69]=g_hwnd; windowCount=70;
    assert(IsWindowActuallyVisible(g_hwnd)); // Bounded hook capacity fails awake.
    assert(g_visibilityLocationHookCount<=64);
}
static void callbacks(void) {
    init(); UpdateJsVisibilityState(g_hwnd);
    assert(g_webViewSuspendPending);
    UINT64 old=g_suspendRequestId;
    ActivateMainWebView();
    assert(g_suspendRequestId!=old && !g_webViewSuspendPending);
    DeactivateMainWebView(); assert(pendingCount==2 && g_webViewSuspendPending);
    finish_suspend(0,S_OK,TRUE); // Cancelled callback after a new suspend.
    assert(g_webViewSuspendPending && !g_webViewSuspended);
    finish_suspend(1,S_OK,TRUE);
    assert(g_webViewSuspended && !g_webViewSuspendPending);
    ActivateMainWebView(); DeactivateMainWebView();
    ++g_suspendRequestId; // WebView replacement invalidates old COM completions.
    g_webViewSuspendPending=FALSE;
    finish_suspend(2,S_OK,TRUE);
    assert(!g_webViewSuspended && !g_webViewSuspendPending);

    LivenessPingHandler ping={.requestId=1}; g_livenessRequestId=2; g_webViewPingOutstanding=TRUE;
    LivenessPingHandler_Invoke((ICoreWebView2ExecuteScriptCompletedHandler *)&ping,S_OK,L"1");
    assert(g_webViewPingOutstanding);
    ping.requestId=2;
    LivenessPingHandler_Invoke((ICoreWebView2ExecuteScriptCompletedHandler *)&ping,-1,NULL);
    assert(g_webViewPingOutstanding);
    LivenessPingHandler_Invoke((ICoreWebView2ExecuteScriptCompletedHandler *)&ping,S_OK,L"1");
    assert(!g_webViewPingOutstanding);
}
static void navigation(void) {
    init(); windows[0].visible=FALSE; StopVisibilityTracking(g_hwnd);
    UpdateJsVisibilityState(g_hwnd); finish_suspend(0,S_OK,TRUE);
    unsigned before=resumes;
    ResetTargetPageInBackground(); // Already at target: stay asleep, no completion needed.
    assert(resumes==before && !navigations && g_webViewSuspended);
    wcscpy(currentUrl,L"https://example.com/elsewhere");
    ResetTargetPageInBackground();
    assert(navigations==1 && rendered && !g_webViewSuspended && g_mainNavigationLoading);
    DeactivateMainWebView(); assert(rendered && suspends==1);
    g_mainNavigationLoading=FALSE; OnMainNavigationCompleted();
    assert(g_webViewSettlePending && timers[ID_TIMER_WEBVIEW_PRELOAD]==1500);
    DeactivateMainWebView(); assert(rendered && suspends==1);
    fire_timer(ID_TIMER_WEBVIEW_PRELOAD); // Old WM_TIMER queued before the new completion.
    assert(g_webViewSettlePending && rendered && suspends==1);
    expire_timer(ID_TIMER_WEBVIEW_PRELOAD);
    assert(!g_webViewSettlePending && !rendered && suspends==2);
    finish_suspend(1,S_OK,TRUE);

    PrepareMainWebViewNavigation(); g_mainNavigationLoading=FALSE; OnMainNavigationCompleted();
    PrepareMainWebViewNavigation(); // New load cancels the previous settle deadline.
    assert(!timers[ID_TIMER_WEBVIEW_PRELOAD]);
    fire_timer(ID_TIMER_WEBVIEW_PRELOAD); // Even a stale timeout cannot freeze the load.
    assert(rendered && g_mainNavigationLoading && suspends==2);
    g_mainNavigationLoading=FALSE; OnMainNavigationCompleted();
    windows[0].visible=TRUE; foreground=g_hwnd; StartVisibilityTracking(g_hwnd);
    UpdateJsVisibilityState(g_hwnd);
    assert(!g_webViewSettlePending && rendered && !timers[ID_TIMER_WEBVIEW_PRELOAD]);
}
static void recovery(void) {
    init(); windows[0].visible=FALSE; StopVisibilityTracking(g_hwnd);
    UpdateJsVisibilityState(g_hwnd); finish_suspend(0,S_OK,TRUE);
    PrewarmMainWebView(); assert(rendered && g_webViewPrewarmActive);
    unsigned before=timerSets;
    for (int i=0;i<10000;++i) { ++ticks; PrewarmMainWebView(); }
    assert(timerSets==before); // Hover traffic only updates a timestamp.
    ticks+=1000; fire_timer(ID_TIMER_WEBVIEW_PREWARM);
    assert(g_webViewPrewarmActive && timers[ID_TIMER_WEBVIEW_PREWARM]==59000);
    ticks+=60000; fire_timer(ID_TIMER_WEBVIEW_PREWARM);
    assert(!g_webViewPrewarmActive && !rendered);
    finish_suspend(1,S_OK,TRUE);
    resumeFails=TRUE; ActivateMainWebView();
    assert(timers[ID_TIMER_WEBVIEW_RESUME_RETRY]==1000);
    fire_timer(ID_TIMER_WEBVIEW_RESUME_RETRY); fire_timer(ID_TIMER_WEBVIEW_RESUME_RETRY);
    assert(!timers[ID_TIMER_WEBVIEW_RESUME_RETRY] && posts==1);
    assert(lastPost==WM_APP_WEBVIEW_RECREATE);
    DeactivateMainWebView(); assert(!timers[ID_TIMER_WEBVIEW_RESUME_RETRY]);
    resumeFails=FALSE; visibilityFails=TRUE;
    g_webViewSuspended=FALSE; g_webViewSuspendPending=FALSE; g_controllerVisible=TRUE;
    unsigned beforeSuspend=suspends;
    DeactivateMainWebView(); assert(suspends==beforeSuspend); // Never suspend a visible controller.
}
static void health(void) {
    init(); foreground=g_hwnd;
    g_mainNavigationLoading=TRUE; ArmMainHealthCheck();
    assert(!scripts && !timers[ID_TIMER_HEALTH_CHECK]);
    g_mainNavigationLoading=FALSE; OnMainNavigationCompleted();
    assert(scripts==1 && g_healthCheckPending && timers[ID_TIMER_HEALTH_CHECK]==3000);
    fire_timer(ID_TIMER_HEALTH_CHECK); // A stale message cannot shorten the new deadline.
    assert(g_healthCheckPending && !kicks && !rebuilds);
    UINT64 old=g_frameProbeId;
    for (int i=0;i<1000;++i) ArmMainHealthCheck();
    assert(scripts==1 && old==g_frameProbeId);
    g_framePongSeen=TRUE; expire_timer(ID_TIMER_HEALTH_CHECK);
    assert(!timers[ID_TIMER_HEALTH_CHECK] && !rebuilds);
    g_presentationUnverified=TRUE; uniformPage=TRUE; ArmMainHealthCheck();
    g_framePongSeen=TRUE; expire_timer(ID_TIMER_HEALTH_CHECK);
    assert(kicks==1 && !rebuilds && g_healthCheckPending);
    g_framePongSeen=TRUE; expire_timer(ID_TIMER_HEALTH_CHECK);
    assert(!rebuilds && !g_presentationUnverified && !g_healthCheckPending);
    g_healthKicked=FALSE; uniformPage=FALSE;
    ArmMainHealthCheck(); expire_timer(ID_TIMER_HEALTH_CHECK);
    assert(kicks==2 && !rebuilds);
    expire_timer(ID_TIMER_HEALTH_CHECK); assert(rebuilds==1);
    ArmMainHealthCheck(); expire_timer(ID_TIMER_HEALTH_CHECK); assert(rebuilds==1);
    assert(!timers[ID_TIMER_HEALTH_CHECK]);
    g_healthHealed=FALSE; g_healthKicked=FALSE; foreground=NULL;
    ArmMainHealthCheck(); expire_timer(ID_TIMER_HEALTH_CHECK);
    assert(rebuilds==1 && kicks==2); // Non-foreground frame throttling is not a dead renderer.
}
static void load_fails_again(void) {
    // Production's completion keeps g_mainLoadFailed set for another failure.
    g_mainNavigationLoading=FALSE; OnMainNavigationCompleted();
    expire_timer(ID_TIMER_WEBVIEW_PRELOAD);
}
static void retry(void) {
    init(); foreground=g_hwnd;
    assert(IsConnectivityLoadFailure(COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN,0));
    assert(IsConnectivityLoadFailure(COREWEBVIEW2_WEB_ERROR_STATUS_DISCONNECTED,0));
    assert(IsConnectivityLoadFailure(COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN,502));
    assert(!IsConnectivityLoadFailure(COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN,404));
    assert(!IsConnectivityLoadFailure(COREWEBVIEW2_WEB_ERROR_STATUS_UNKNOWN,200));
    assert(!IsConnectivityLoadFailure(COREWEBVIEW2_WEB_ERROR_STATUS_CERTIFICATE_IS_INVALID,0));
    assert(!IsConnectivityLoadFailure(COREWEBVIEW2_WEB_ERROR_STATUS_OPERATION_CANCELED,0));

    // A burst of route changes queues one message and parameter refreshes
    // none; every message restarts the one-shot quiet period.
    unsigned before=posts, sets=timerSets;
    OnNetworkRouteChange(NULL,NULL,MibParameterNotification); assert(posts==before);
    for (int i=0;i<1000;++i) OnNetworkRouteChange(NULL,NULL,MibAddInstance);
    assert(posts==before+1 && lastPost==WM_APP_NETWORK_CHANGED);
    network_message();
    assert(!g_networkChangePosted && timers[ID_TIMER_NETWORK_SETTLE]==1500 && timerSets==sets+1);
    OnNetworkRouteChange(NULL,NULL,MibDeleteInstance); assert(posts==before+2);
    network_message(); assert(timerSets==sets+2 && !cooldownExpiries && !navigations);

    // Nothing failed: a network change only lets the proxy re-probe.
    expire_timer(ID_TIMER_NETWORK_SETTLE); assert(!timers[ID_TIMER_NETWORK_SETTLE]);
    assert(cooldownExpiries==1 && !navigations);
    UpdateJsVisibilityState(g_hwnd); assert(!navigations);

    // A failed page is fetched again from the URL it failed on, not reloaded
    // or sent home, and a load in progress is never restarted.
    wcscpy(currentUrl,L"https://example.com/deep"); g_mainLoadFailed=TRUE;
    fire_timer(ID_TIMER_NETWORK_SETTLE);
    assert(navigations==1 && g_mainNavigationLoading && rendered);
    assert(!wcscmp(currentUrl,L"https://example.com/deep"));
    fire_timer(ID_TIMER_NETWORK_SETTLE); UpdateJsVisibilityState(g_hwnd);
    assert(navigations==1);

    // Failing again, it waits for a trigger: staying on screen is not one,
    // returning to it is.
    load_fails_again();
    UpdateJsVisibilityState(g_hwnd); assert(navigations==1);
    foreground=NULL; UpdateJsVisibilityState(g_hwnd);
    assert(g_jsVisibility==JS_VISIBILITY_HIDDEN && navigations==1);
    finish_suspend(pendingCount-1,S_OK,TRUE); assert(g_webViewSuspended);
    foreground=g_hwnd; UpdateJsVisibilityState(g_hwnd);
    assert(navigations==2 && g_mainNavigationLoading && !g_webViewSuspended);

    // Hidden and asleep: a network change retries in the background, and a
    // page that then loads settles back to sleep with nothing left to retry.
    load_fails_again();
    foreground=NULL; UpdateJsVisibilityState(g_hwnd);
    finish_suspend(pendingCount-1,S_OK,TRUE); assert(g_webViewSuspended);
    fire_timer(ID_TIMER_NETWORK_SETTLE);
    assert(navigations==3 && rendered && !g_webViewSuspended);
    g_mainLoadFailed=FALSE; load_fails_again();
    assert(!rendered && g_webViewSuspendPending);
    fire_timer(ID_TIMER_NETWORK_SETTLE); assert(navigations==3);

    // Home over an error page for the target itself loads it again, once.
    wcscpy(currentUrl,g_initialUrl); g_mainLoadFailed=TRUE;
    ResetTargetPageIfNeeded(); assert(navigations==4 && g_mainNavigationLoading);
    ResetTargetPageIfNeeded(); assert(navigations==4);
    g_mainNavigationLoading=FALSE; g_mainLoadFailed=FALSE;
    ResetTargetPageIfNeeded(); assert(navigations==4);

    // An unreadable URL falls back to the configured target.
    g_mainNavigationLoading=FALSE; g_mainLoadFailed=TRUE; sourceFails=TRUE;
    fire_timer(ID_TIMER_NETWORK_SETTLE); assert(targetReloads==1 && navigations==4);
    sourceFails=FALSE;
}
int main(int argc, char **argv) {
    assert(argc==2);
    struct { const char *name; void (*fn)(void); } cases[]={
        {"coverage",coverage},{"geometry",geometry},{"events",events},{"lifecycle",lifecycle},
        {"callbacks",callbacks},{"navigation",navigation},{"recovery",recovery},{"health",health},
        {"retry",retry}
    };
    BOOL found=FALSE;
    for (size_t i=0;i<sizeof(cases)/sizeof(cases[0]);++i) if (!strcmp(argv[1],cases[i].name)) {
        found=TRUE; cases[i].fn();
    }
    assert(found && !liveRegions);
    StopVisibilityTracking(g_hwnd);
    for (size_t i=0;i<pendingCount;++i) if (pending[i]) pending[i]->lpVtbl->Release(pending[i]);
    return 0;
}
