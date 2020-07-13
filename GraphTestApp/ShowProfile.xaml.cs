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

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=234238

namespace GraphTestApp
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class ShowProfilePage : Page
    {
        public ShowProfilePage()
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

        protected override async void OnNavigatedTo(NavigationEventArgs e)
        {
            // Get the Graph client from the provider
            var graphClient = ProviderManager.Instance.GlobalProvider.Graph;

            try
            {
                // Get the events
                
                var Manager = await graphClient.Me.Manager.Request()
                    .GetAsync();
                var myPhoto = await graphClient.Me.Photo.Content.Request()
                    .GetAsync();
                var me = await graphClient.Me.Request()
                    .GetAsync();          

                BitmapImage bitmap = new BitmapImage();
                await bitmap.SetSourceAsync(myPhoto.AsRandomAccessStream());
                 Photo.Source = bitmap;

                PlaceName.Text = me.DisplayName;
                PlaceTitle.Text = me.JobTitle;
                PlaceLocation.Text = me.OfficeLocation;
                PlaceEmail.Text = me.Mail;
                PlaceManager.Text = ((Microsoft.Graph.User)Manager).DisplayName;
                
                
            }
            catch (Microsoft.Graph.ServiceException ex)
            {
                ShowNotification($"Exception getting events: {ex.Message}");
            }

            base.OnNavigatedTo(e);
        }
    }
}
