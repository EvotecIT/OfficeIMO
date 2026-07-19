namespace OfficeIMO.Email;

/// <summary>Behavior used when an Outlook voting option is selected.</summary>
public enum OutlookVoteSendBehavior {
    /// <summary>Send the response automatically.</summary>
    Automatic = 1,
    /// <summary>Prompt before sending or editing the response.</summary>
    Prompt = 2
}

/// <summary>One option decoded from an Outlook PidLidVerbStream value.</summary>
public sealed class OutlookVoteOption {
    /// <summary>Creates a voting option.</summary>
    public OutlookVoteOption(string displayName) {
        if (displayName == null) throw new ArgumentNullException(nameof(displayName));
        if (displayName.Length == 0) throw new ArgumentException("A voting option requires a display name.", nameof(displayName));
        DisplayName = displayName;
    }

    /// <summary>Localized display name stored in the verb stream.</summary>
    public string DisplayName { get; set; }
    /// <summary>Monotonically increasing option identifier.</summary>
    public int Id { get; set; }
    /// <summary>Response-send behavior.</summary>
    public OutlookVoteSendBehavior SendBehavior { get; set; } = OutlookVoteSendBehavior.Prompt;
    /// <summary>Whether the response should use U.S.-style reply headers.</summary>
    public bool UseUsReplyHeaders { get; set; }
}

/// <summary>Outlook message voting options and selected response.</summary>
public sealed class OutlookVoting {
    private readonly List<OutlookVoteOption> _options = new List<OutlookVoteOption>();

    /// <summary>Decoded voting options.</summary>
    public IList<OutlookVoteOption> Options => _options;
    /// <summary>Voting option selected on a response message.</summary>
    public string? Response { get; set; }
    /// <summary>Original PidLidVerbStream bytes, retained even when decoding fails.</summary>
    public byte[]? RawVerbStream { get; internal set; }
    /// <summary>Whether <see cref="Options"/> was decoded successfully from <see cref="RawVerbStream"/>.</summary>
    public bool OptionsDecoded { get; internal set; }
    internal bool OptionsClearRequested { get; private set; }

    /// <summary>Removes all voting options, including any retained undecodable verb stream.</summary>
    public void ClearOptions() {
        _options.Clear();
        RawVerbStream = null;
        OptionsDecoded = true;
        OptionsClearRequested = true;
    }

    internal void ResetProjectionState() => OptionsClearRequested = false;
}
