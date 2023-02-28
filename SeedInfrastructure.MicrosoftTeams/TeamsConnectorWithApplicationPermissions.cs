using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace SeedInfrastructure.MicrosoftTeams
{
    public static class TeamsConnectorWithApplicationPermissions
    {
        public static async Task<Event> CreateOnlineMeeting(GraphServiceClient graphClient, string userId, Event _event)
        {
            var createdOnlineMeeting = await graphClient.Users[userId].Events
                  .Request()
                  .Header("Prefer", "outlook.timezone=\"UTC\"")
                  .AddAsync(_event);

            return createdOnlineMeeting;
        }



        public static async Task<List<Event>> GetOnlineMeetings(GraphServiceClient graphClient, string userId)
        {
            var @events = await graphClient.Users[userId].Events
                .Request()
                .Select("subject,body,bodyPreview,organizer,attendees,start,end,location,locations")
                .GetAsync();

            return @events.ToList();
        }



        public static async Task<Event> GetOnlineMeetingById(GraphServiceClient graphClient, string userId, string eventId)
        {
            var @event = await graphClient.Users[userId].Events[eventId]
                .Request()
                .Select("subject,body,bodyPreview,organizer,attendees,start,end,location,locations")
                .GetAsync();

            return @event;
        }

        public static async Task UpdateOnlineMeeting(GraphServiceClient graphClient, string userId, Event _event, string eventId)
        {
            await graphClient.Users[userId].Events[eventId]
                .Request()
                .UpdateAsync(_event);

        }


        public static async Task DeleteOnlineMeeting(GraphServiceClient graphClient, string userId, string eventId)
        {
            await graphClient.Users[userId].Events[eventId]
                .Request()
                .DeleteAsync();
        }





        private static async Task<string> CreateTeam2(GraphServiceClient graphClient, Team team)
        {
            /*             await graphClient.Teams
                     .Request()
                     .AddAsync(team);

           return null;*/

            string location;
            BaseRequest request = (BaseRequest)graphClient.Teams.Request();
            request.ContentType = "application/json";
            request.Method = HttpMethods.POST;

            using (HttpResponseMessage response = await request.SendRequestAsync(team, CancellationToken.None))
                location = response.Headers.Location.ToString();


            string[] locationParts = location.Split(new[] { '\'', '/', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
            string teamId = locationParts[1];
            string operationId = locationParts[3];

            // before querying the first time we must wait some secs, else we get a 404
            // int delayInMilliseconds = 1_000;
            int i = 1;
            while (i <= 10)
            {

                await Task.Delay(1000);
                // lets see how far the teams creation process is
                TeamsAsyncOperation operation = await graphClient.Teams[teamId].Operations[operationId].Request().GetAsync();
                if (operation.Status == TeamsAsyncOperationStatus.Succeeded)
                    break;

                if (operation.Status == TeamsAsyncOperationStatus.Failed)
                    throw new Exception($"Failed to create team '{team.DisplayName}': {operation.Error.Message} ({operation.Error.Code})");

                // according to the docs, we should wait > 30 secs between calls
                // https://learn.microsoft.com/en-us/graph/api/resources/teamsasyncoperation?view=graph-rest-1.0
                //  delayInMilliseconds = 30_000;
                i++;
            }


            return teamId;

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


        public static Team CreateTeamEntity(string userId, string displayName, string description)
        {
            var team = new Team
            {
                DisplayName = displayName,
                Description = description,
                Members = new TeamMembersCollectionPage()
                {
                    new AadUserConversationMember
                    {
                        Roles = new List<String>()
                        {
                            "owner"
                        },
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{userId}')"}
                        }
                    }
                },
                AdditionalData = new Dictionary<string, object>()
                {
                    {"template@odata.bind", "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"}
                }
            };

            return team;
        }


        public static async Task<List<Team>> GetTeams(GraphServiceClient graphClient, string userId)
        {
            var joinedTeams = await graphClient.Users[userId].JoinedTeams
                .Request()
                .GetAsync();
            return joinedTeams.ToList();
        }


        public static async Task<string> CreateTeam(GraphServiceClient graphClient, Team team)
        {

            string location;
            var request = graphClient.Teams.Request();
            request.ContentType = "application/json";
            request.Method = HttpMethods.POST;

            var response = await request.AddResponseAsync(team);

            var uu = response.GetResponseObjectAsync();


            string[] locationParts = response.HttpHeaders.Location.ToString().Split(new[] { '\'', '/', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
            string teamId = locationParts[1];
            string operationId = locationParts[3];



            //  Console.WriteLine($"The location is: {teamId}");
            // var response = await request.AddAsync(team);
            //   await Task.Delay(3000);

            //while (uu.Status != TaskStatus.RanToCompletion)
            //{
            //    Console.WriteLine("waiting");
            //}

            // var t = uu.Result;
            //  Console.WriteLine($"The Id is: {response.Content.Headers.GetValues("location")}");
            return teamId;
        }

        public static async Task AddMembersToTeam(GraphServiceClient graphClient, string teamId, List<ConversationMember> members)
        {
            await graphClient.Teams[teamId].Members
                .Add(members)
                .Request()
                .PostAsync();

        }

        public static async Task AddMembersToTeam(GraphServiceClient graphClient, string teamId, IEnumerable<Attendee> attendees)
        {
            var members = new List<ConversationMember>();
            foreach (var meetingAttendee in attendees)
            {
                members.Add(
                    new AadUserConversationMember
                    {
                        Roles = new List<String>()
                        {

                        },
                        AdditionalData = new Dictionary<string, object>()
                        {
                            {"user@odata.bind", $"https://graph.microsoft.com/v1.0/users('{meetingAttendee.EmailAddress.Address}')"}
                        }
                    });
            }
            await graphClient.Teams[teamId].Members
                .Add(members)
                .Request()
                .PostAsync();

        }


        public static async Task DeleteTeam(GraphServiceClient graphClient, string teamId)
        {
            await graphClient.Groups[teamId]
                .Request()
                .DeleteAsync();
        }





        public static async Task<Group> CreateGroup(GraphServiceClient graphClient, string userId, string displayName, string description, string mailNickName, IEnumerable<Attendee> attendees)
        {

            var additionalData = new Dictionary<string, object>()
            {
                {"owners@odata.bind", new List<string>()},
                {"members@odata.bind", new List<string>()}
            };

            (additionalData["owners@odata.bind"] as List<string>)?.Add($"https://graph.microsoft.com/v1.0/users/{userId}");


            foreach (var attendee in attendees)
            {
                (additionalData["members@odata.bind"] as List<string>)?.Add($"https://graph.microsoft.com/v1.0/users('{attendee.EmailAddress.Address}')");


            }



            var group = new Group
            {
                Description = description,
                DisplayName = displayName,
                GroupTypes = new List<String>()
                {
                    "Unified"
                },
                MailEnabled = false,
                MailNickname = mailNickName,
                SecurityEnabled = true,
                AdditionalData = additionalData
            };
            /*
            var grp = await graphClient.Groups
                .Request()
                .AddAsync(group);
            */




            var request = graphClient.Groups.Request();
            request.ContentType = "application/json";
            request.Method = HttpMethods.POST;

            var response = await request.AddResponseAsync(group);

            var teamToBeCreated = response.GetResponseObjectAsync();



            return teamToBeCreated.Result;
        }


        public static async Task<Team> CreateTeamFromGroup(GraphServiceClient graphClient, Group group)
        {
            //    await Task.Delay(2000);
            //    Thread.Sleep(2000);
            var team = new Team
            {
                MemberSettings = new TeamMemberSettings
                {
                    AllowCreatePrivateChannels = true,
                    AllowCreateUpdateChannels = true
                },
                MessagingSettings = new TeamMessagingSettings
                {
                    AllowUserEditMessages = true,
                    AllowUserDeleteMessages = true
                },
                FunSettings = new TeamFunSettings
                {
                    AllowGiphy = true,
                    GiphyContentRating = GiphyRatingType.Strict
                }
            };

            var createdTeam = await graphClient.Groups[group.Id].Team
                .Request()
                .PutAsync(team);

            //   var y = createdTeam.DisplayName;

            return createdTeam;
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
