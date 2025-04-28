using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
/// <FLOWERBOX file="EventsExceptionsUpdater.cs">
/// <Created_By>
/// COMET DEV TEAM
/// </Created_By>
/// <Purpose>
/// Update events and exceptions in the code.
/// </Purpose>
/// <Revise_History>
/// 4/27/2025 - Initial release
/// </Revise_History>
/// </FLOWERBOX>

/// <NAMESPACE name="COMET.Resources">
/// <Purpose>
///
/// </Purpose>
/// </NAMESPACE>
namespace COMET.Resources
{
   /// <CLASS name="EventsUpdater">
   /// <Purpose>
   /// Class for EventsUpdater
   /// </Purpose>
   /// </CLASS>
    public class EventsUpdater
    {
       /// <CLASS name="Event">
       /// <Purpose>
       /// Class for Event
       /// </Purpose>
       /// </CLASS>
        public class Event
        {
           /// <PROPERTY name="Name">
           /// <Purpose>
           ///  Name of Event
           /// </Purpose>
           /// </PROPERTY>
            public string Name { get; set; }

            /// <PROPERTY name="Description">
            /// <Purpose>
            ///  Description of Event
            /// </Purpose>
            /// </PROPERTY>
            public List<string> Description { get; set; } = new List<string>();
        }

       /// <METHOD name="UpdateEvents">
       /// <Purpose> 
       ///  Updates the comment for events in the code.
       /// </Purpose>
       /// <Parameters>
       ///    currentEvents(string):
       ///    updatedEvents(List<string>):
       ///    indent(string):
       /// </Parameters>
       /// </METHOD>
        public static string UpdateEvents(string currentEvents, List<string> updatedEvents, string indent)
        {
            List<Event> ParseCurrentEvents(string eventsText)
            {
                var events = new List<Event>();
                if (string.IsNullOrWhiteSpace(eventsText))
                    return events;

                // Split by '///' and remove empty entries
                var entries = eventsText.Split(new[] { "///" }, StringSplitOptions.RemoveEmptyEntries);
                Event currentEvent = null;

                foreach (var entry in entries)
                {
                    var trimmedEntry = entry.Trim();
                    if (string.IsNullOrEmpty(trimmedEntry))
                        continue;

                    // Check if the entry is an event (matches name.event: or word: pattern)
                    var match = Regex.Match(trimmedEntry, @"^((\w+\.\w+)|\w+):(?:\s*(.*))?$");
                    if (match.Success)
                    {
                        if (currentEvent != null)
                            events.Add(currentEvent);

                        currentEvent = new Event
                        {
                            // Use group 2 (name.event) if matched, otherwise group 1 (word)
                            Name = match.Groups[2].Success ? match.Groups[2].Value : match.Groups[1].Value
                        };
                        if (!string.IsNullOrEmpty(match.Groups[3].Value))
                            currentEvent.Description.Add(match.Groups[3].Value);
                    }
                    else if (currentEvent != null)
                    {
                        // Non-event entry is a description line
                        currentEvent.Description.Add(trimmedEntry);
                    }
                }

                if (currentEvent != null)
                    events.Add(currentEvent);

                return events;
            }

            List<Event> ParseUpdatedEvents(List<string> eventsText)
            {
                var events = new List<Event>();
                if (eventsText == null || !eventsText.Any())
                    return events;

                foreach (var entry in eventsText)
                {
                    var trimmedEntry = entry?.Trim();
                    if (string.IsNullOrEmpty(trimmedEntry))
                        continue;

                    // Assume all entries are events (matches name.event or word pattern)
                    var match = Regex.Match(trimmedEntry, @"^((\w+\.\w+)|\w+)$");
                    if (match.Success)
                    {
                        events.Add(new Event
                        {
                            // Use group 2 (name.event) if matched, otherwise group 1 (word)
                            Name = match.Groups[2].Success ? match.Groups[2].Value : match.Groups[1].Value,
                            Description = new List<string>()
                        });
                    }
                }

                return events;
            }

            string BuildEvents(List<Event> events, List<Event> referenceEvents, string indents)
            {
                var sb = new StringBuilder();
                for (int i = 0; i < events.Count; i++)
                {
                    var evt = events[i];
                    var desc = (referenceEvents != null && i < referenceEvents.Count) ? referenceEvents[i].Description : evt.Description;
                    sb.Append($"{indents}///    {evt.Name}:");
                    if (desc.Count > 0)
                        sb.Append($" {desc[0]}");
                    sb.AppendLine(); // Add newline after event line

                    // Add additional description lines
                    for (int j = 1; j < desc.Count; j++)
                        sb.AppendLine($"{indents}///    {desc[j]}");
                }
                return sb.ToString().TrimEnd('\n');
            }

            var currentEventList = ParseCurrentEvents(currentEvents);
            var updatedEventList = ParseUpdatedEvents(updatedEvents);

            if (updatedEventList.Count <= currentEventList.Count)
            {
                // Update names, transfer descriptions only for matching event names
                foreach (var updatedEvent in updatedEventList)
                {
                    // Find a current event with the same name
                    var matchingCurrentEvent = currentEventList.FirstOrDefault(e => e.Name == updatedEvent.Name);
                    if (matchingCurrentEvent != null)
                    {
                        // Transfer the description from the matching current event
                        updatedEvent.Description = matchingCurrentEvent.Description;
                    }
                    // If no match, updatedEvent.Description remains as is (empty)
                }
            }
            else
            {
                // Copy descriptions for matching events, leave new ones as is
                foreach (var updatedEvent in updatedEventList)
                {
                    // Find a current event with the same name
                    var matchingCurrentEvent = currentEventList.FirstOrDefault(e => e.Name == updatedEvent.Name);
                    if (matchingCurrentEvent != null)
                    {
                        // Transfer the description from the matching current event
                        updatedEvent.Description = matchingCurrentEvent.Description;
                    }
                    // New events or non-matching names keep their descriptions (empty)
                }
            }

            return $"{BuildEvents(updatedEventList, null, indent)}";
        }
    }
}