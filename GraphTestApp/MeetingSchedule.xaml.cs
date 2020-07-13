using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;
using Microsoft.Toolkit.Graph.Providers;
using Microsoft.Toolkit.Uwp.UI.Controls;
using Newtonsoft.Json;
using Windows.UI.Xaml.Media.Imaging;
using Microsoft.Graph;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=234238

namespace GraphTestApp
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class ShowMeetingSchedulePage : Page
    {
        public ShowMeetingSchedulePage()
        {
            this.InitializeComponent();
        }

        private void ShowNotification(string message)
        {
            // Get the main page that contains the InAppNotification
            var mainPage = (Window.Current.Content as Frame).Content as MainPage;

            // Get the notification control
            var notification = mainPage.FindName("Notification") as InAppNotification;

            notification.Show(message);
        }

        private async void CalendarInviteButton_Click(object sender, RoutedEventArgs e)
        {
            // Call app specific code to subscribe to the service. For example:
            //BigTextArea.Text = PersonName.Text;
            try
            {
                GraphServiceClient graphClient = ProviderManager.Instance.GlobalProvider.Graph;
                var user = await graphClient.Users[PersonEmail.Text.ToString()]
                .Request()
                .GetAsync();

                var attendees = new List<AttendeeBase>()
            {
                new AttendeeBase
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = user.Mail,
                        Name = user.DisplayName
                    },
                    Type = AttendeeType.Required
                }
            };

                var timeConstraint = new TimeConstraint
                {
                    TimeSlots = new List<TimeSlot>()
                {
                    new TimeSlot
                    {
                        Start = new DateTimeTimeZone
                        {
                            DateTime = "2020-07-12T19:58:00.557Z",
                            TimeZone = "Pacific Standard Time"
                        },
                        End = new DateTimeTimeZone
                        {
                            DateTime = "2020-07-19T19:58:00.557Z",
                            TimeZone = "Pacific Standard Time"
                        }
                    }
                }
                };

                var locationConstraint = new LocationConstraint
                {
                    IsRequired = false,
                    SuggestLocation = true,
                    Locations = new List<LocationConstraintItem>()
                {
                    new LocationConstraintItem
                    {
                        DisplayName = "Conf Room 32/1368",
                        LocationEmailAddress = "conf32room1368@imgeek.onmicrosoft.com"
                    }
                }
                };

                var meetingDuration = new Microsoft.Graph.Duration("PT1H");

                MeetingTimeSuggestionsResult response = await graphClient.Me
                     .FindMeetingTimes(attendees, null, null, meetingDuration, null, null, null, null)
                     //.FindMeetingTimes(attendees, locationConstraint, timeConstraint, meetingDuration, null, null, null, null)
                     .Request()
                     .PostAsync();

                if (response.MeetingTimeSuggestions.Count() == 0)
                {
                    BigTextArea.Text = "No common meeting time found";
                }
                else
                {
                    var StartTime = response.MeetingTimeSuggestions.First().MeetingTimeSlot.Start;
                    var EndTime = response.MeetingTimeSuggestions.First().MeetingTimeSlot.End;

                    // Calendar Invite
                    var @event = new Event
                    {
                        Subject = "My Calendar Invite via code",
                        Start = StartTime,
                        End = EndTime,
                        Attendees = new List<Attendee>()
                            {
                                new Attendee
                                {
                                    EmailAddress = new EmailAddress
                                    {
                                        Address = user.Mail,
                                        Name = user.DisplayName
                                    },
                                    Type = AttendeeType.Required
                                }
                            },
                        IsOnlineMeeting = true,
                        OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness
                    };

                    await graphClient.Me.Events
                        .Request()
                        .AddAsync(@event);

                    var dateconvertinstance = new GraphDateTimeTimeZoneConverter();
                    BigTextArea.Text = "Calendar invite sent from " + dateconvertinstance.Convert(StartTime,null,null,null) + " to " + dateconvertinstance.Convert(EndTime, null, null, null) + "\n\n" +
                        "Please check your calendar for teams meeting.";

                }
            }
            catch (Microsoft.Graph.ServiceException ex)
            {
                if (ex.Error.Code == "Request_ResourceNotFound")
                    BigTextArea.Text = "This user does not exit in the directory";
                else
                    BigTextArea.Text = "Error message - \n" + ex.Message;
            }
        }



        /*
        protected override async void OnNavigatedTo(NavigationEventArgs e)
        {
            // Get the Graph client from the provider
            var graphClient = ProviderManager.Instance.GlobalProvider.Graph;

            try
            {
                // Get the events
                var events = await graphClient.Me.Events.Request()
                    .Select("subject,organizer,start,end")
                    .OrderBy("createdDateTime DESC")
                    .GetAsync();

                EventList.ItemsSource = events.CurrentPage.ToList();
            }
            catch (Microsoft.Graph.ServiceException ex)
            {
                ShowNotification($"Exception getting events: {ex.Message}");
            }

            base.OnNavigatedTo(e);
        }
        */
    }
}
