using Microsoft.Extensions.Configuration;
using RegistrationSummary.Common.Configurations;
using RegistrationSummary.Common.Interfaces;
using RegistrationSummary.Maui.Services;
using RegistrationSummary.Maui.ViewModels;
using RegistrationSummary.Maui.Views;
using System.Reflection;
using CommunityToolkit.Maui;

namespace RegistrationSummary.Maui;

public static class MauiProgram
{
    public static string EventsDataFilePath { get; set; }
    public static string CredentialsFilePath { get; set; }

    public static MauiApp CreateMauiApp()
    {
        var builder = MauiApp.CreateBuilder();
        builder
            .UseMauiApp<App>()
            .ConfigureFonts(fonts =>
            {
                fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
                fonts.AddFont("OpenSans-Semibold.ttf", "OpenSansSemibold");
            });

        builder.UseMauiApp<App>().UseMauiCommunityToolkit();

        FileService.CopyTemplateFilesIfDontExist();

        IConfiguration configuration = BuildConfiguration();
        ConfigureServices(builder.Services, configuration);

        return builder.Build();
    }

    private static IConfiguration BuildConfiguration()
    {
        var assembly = Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream($"{assembly.GetName().Name}.appsettings.json");

        if (stream == null)
            return null;

        var configBuilder = new ConfigurationBuilder()
            .AddJsonStream(stream);

        return configBuilder.Build();
    }

    private static void ConfigureServices(IServiceCollection services, IConfiguration configuration)
    {
        MailerConfiguration mailerConfiguration = new MailerConfiguration();
        var settings = FileService.GetSettings();

        mailerConfiguration.Mail = settings.Mail;
        mailerConfiguration.Password = settings.Password;
        mailerConfiguration.ServerName = settings.ServerName;
        mailerConfiguration.ServerPort = settings.ServerPort;

        EventsDataFilePath = settings.EventsDataFilePath;
        CredentialsFilePath = settings.CredentialsFilePath;
        var emailsTemplatesFilePath = settings.EmailsTemplatesFilePath;
        var rawDataTabName = settings.RawDataTabName;
        var preprocessedDataTabName = settings.PreprocessedDataTabName;
        var groupBalanceTabName = settings.GroupBalanceTabName;
        var summaryTabName = settings.SummaryTabName;
        var leaderText = settings.LeaderText;
        var followerText = settings.FollowerText;
        var soloText = settings.SoloText;
        var testMailRecepientEmailAddress = settings.TestMailRecepientEmailAddress;
        var prices = settings.Prices;

        if (
            string.IsNullOrEmpty(EventsDataFilePath) ||
            string.IsNullOrEmpty(emailsTemplatesFilePath) ||
            string.IsNullOrEmpty(rawDataTabName) ||
            string.IsNullOrEmpty(preprocessedDataTabName) ||
            string.IsNullOrEmpty(summaryTabName) ||
            prices == null)
            throw new Exception("Configuration could not be read properly.");

        services.AddSingleton(mailerConfiguration);
        services.AddSingleton<IDialogService, DialogService>();

        services.AddViewModel<MainPageViewModel, MainPageView>();
        services.AddViewModel<EventModificationViewModel, EventModificationView>();

        services.AddSingleton(serviceProvider =>
            new MainPageViewModel(
                serviceProvider.GetRequiredService<MailerConfiguration>(),
                serviceProvider.GetRequiredService<IDialogService>(),
                emailsTemplatesFilePath,
                rawDataTabName,
                preprocessedDataTabName,
                summaryTabName,
                groupBalanceTabName,
                leaderText,
                followerText,
                soloText,
                testMailRecepientEmailAddress,
                prices
            )
        );
    }

    private static void AddViewModel<TViewModel, TView>(this IServiceCollection services)
        where TView : ContentPage, new()
        where TViewModel : class
    {
        services.AddTransient<TViewModel>();
        services.AddTransient(s => new TView() { BindingContext = s.GetRequiredService<TViewModel>() });
    }
}