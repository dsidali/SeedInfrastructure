using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SeedInfrastructure.MicrosoftTeams
{
    public static class TeamsConnectorWithDelegatedPermissions
    {

        public static async Task<Event> CreateOnlineMeeting(GraphServiceClient graphClient, Event _event)
        {
            var createdOnlineMeeting = await graphClient.Me.Events
                  .Request()
                  .Header("Prefer", "outlook.timezone=\"UTC\"")
                  .AddAsync(_event);

            return createdOnlineMeeting;
        }
        public static async Task<List<Event>> GetOnlineMeetings(GraphServiceClient graphClient)
        {
            var @events = await graphClient.Me.Events
                .Request()
                .Select("subject,body,bodyPreview,organizer,attendees,start,end,location,locations")
                .GetAsync();

            return @events.ToList();
        }



        public static async Task<Event> GetOnlineMeetingById(GraphServiceClient graphClient, string eventId)
        {
            var @event = await graphClient.Me.Events[eventId]
                .Request()
                .Select("subject,body,bodyPreview,organizer,attendees,start,end,location,locations")
                .GetAsync();

            return @event;
        }

        public static async Task UpdateOnlineMeeting(GraphServiceClient graphClient, Event _event, string eventId)
        {
            await graphClient.Me.Events[eventId]
                .Request()
                .UpdateAsync(_event);

        }


        public static async Task DeleteOnlineMeeting(GraphServiceClient graphClient, string eventId)
        {
            await graphClient.Me.Events[eventId]
                .Request()
                .DeleteAsync();
        }






        public static async Task CreateTeam(GraphServiceClient graphClient, Team team)
        {
            await graphClient.Teams
                     .Request()
                     .AddAsync(team);


        }




        public static Event CreateEventEntity(string subject, string content, DateTime start, DateTime end, string location, List<Attendee> attendees)
        {
            Guid guid = Guid.NewGuid();
            var @event = new Event();

            @event.Subject = subject;
            @event.Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = content
            };
            @event.Start = new DateTimeTimeZone
            {
                DateTime = start.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC"
            };
            @event.End = new DateTimeTimeZone
            {
                DateTime = end.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC"
            };
            @event.Location = new Location
            {
                DisplayName = location
            };

            @event.Attendees = attendees;


            @event.AllowNewTimeProposals = true;
            @event.Importance = Importance.High;
            @event.IsOnlineMeeting = true;
            @event.OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness;
            @event.TransactionId = guid.ToString();
            return @event;
        }

        public static async Task SendTeamMessage(GraphServiceClient graphClient, string teamId, string channelId, string message)
        {

            var chatMessage = new ChatMessage
            {
                Body = new ItemBody
                {
                    Content = message
                }
            };

            await graphClient.Teams[teamId].Channels[channelId].Messages
                .Request()
                .AddAsync(chatMessage);


        }

        public static async Task SendChatMessage(GraphServiceClient graphClient, string chatId, string message)
        {
            var chatMessage = new ChatMessage
            {
                Body = new ItemBody
                {
                    Content = message
                }
            };

            await graphClient.Chats[chatId].Messages
                .Request()
                .AddAsync(chatMessage);
        }



        public static Team CreateTeamEntity(string userId, string displayName, string description)
        {


            var team = new Team
            {
                DisplayName = displayName,
                Description = description,
                AdditionalData = new Dictionary<string, object>()
                {
                    {"template@odata.bind", "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"}
                }
            };


            return team;
        }

        public static List<Attendee> CreateAtteneesList(List<string> list)
        {
            var attendees = new List<Attendee>();
            foreach (var element in list)
            {
                attendees.Add(

                    new Attendee
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = element,
                            //     Name = "Alex Wilber"
                        },
                        Type = AttendeeType.Required
                    }
                );
            }

            return attendees;
        }
    }
}
