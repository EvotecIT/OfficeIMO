using Avalonia.Automation;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioNoticeBarTests {
    [Fact]
    public async Task Notice_shows_actions_only_when_usable_and_announces_failures_assertively() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var notice = new StudioNoticeBar { Message = "The file is open in another program." };
            Assert.False(notice.HasAction);
            Assert.False(notice.IsDismissible);
            Assert.Equal(AutomationLiveSetting.Polite, AutomationProperties.GetLiveSetting(notice));

            notice.ActionText = "Retry";
            Assert.False(notice.HasAction);
            notice.ActionCommand = new RelayCommand(() => { });
            Assert.True(notice.HasAction);
            notice.IsActionVisible = false;
            Assert.False(notice.HasAction);

            notice.DismissCommand = new RelayCommand(() => { });
            Assert.True(notice.IsDismissible);

            notice.Severity = StudioNoticeSeverity.Error;
            Assert.Equal(AutomationLiveSetting.Assertive, AutomationProperties.GetLiveSetting(notice));
            Assert.Equal("The file is open in another program.", AutomationProperties.GetName(notice));
            return true;
        }, CancellationToken.None);
    }
}
