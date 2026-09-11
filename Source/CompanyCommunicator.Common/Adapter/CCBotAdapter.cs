// <copyright file="CCBotAdapter.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Adapter
{
    using System;
    using System.Threading;
    using System.Threading.Tasks;
    using Microsoft.Bot.Builder;
    using Microsoft.Bot.Connector.Authentication;
    using Microsoft.Bot.Schema;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Secrets;

    /// <summary>
    /// Bot framework http adapter instance.
    /// </summary>
    public class CCBotAdapter : CCBotAdapterBase
    {
        private readonly ICertificateProvider certificateProvider;

        /// <summary>
        /// Initializes a new instance of the <see cref="CCBotAdapter"/> class.
        /// </summary>
        /// <param name="botFrameworkAuthentication">credential provider.</param>
        public CCBotAdapter(
            ICertificateProvider certificateProvider,
            BotFrameworkAuthentication botFrameworkAuthentication)
            : base(botFrameworkAuthentication)
        {
            this.certificateProvider = certificateProvider;
        }

        /// <inheritdoc/>
        public override async Task CreateConversationUsingCertificateAsync(string channelId, string serviceUrl, AppCredentials appCredentials, ConversationParameters conversationParameters, BotCallbackHandler callback, CancellationToken cancellationToken)
        {
            var cert = await this.certificateProvider.GetCertificateAsync(appCredentials.MicrosoftAppId);
            var options = new CertificateAppCredentialsOptions()
            {
                AppId = appCredentials.MicrosoftAppId,
                ClientCertificate = cert,
            };

            // NOTE: passing an explicit audience (instead of null) is required for
            // SingleTenant bots on this Bot Framework SDK version (4.21.2). When
            // audience is null, CloudAdapterBase does not always resolve the correct
            // OAuth scope for SingleTenant apps, which results in a token with the
            // wrong audience and a 401 "Authorization has been denied" from the Bot
            // Connector Service, even though the credentials and tenant are correct.
            await this.CreateConversationAsync(appCredentials.MicrosoftAppId, channelId, serviceUrl, AuthenticationConstants.ToChannelFromBotOAuthScope, conversationParameters, callback, cancellationToken);
        }

        /// <inheritdoc/>
        public override async Task CreateConversationUsingSecretAsync(string channelId, string serviceUrl, MicrosoftAppCredentials credentials, ConversationParameters conversationParameters, BotCallbackHandler callback, CancellationToken cancellationToken)
        {
            // See note above regarding explicit audience for SingleTenant bots.
            await this.CreateConversationAsync(credentials.MicrosoftAppId, channelId, serviceUrl, AuthenticationConstants.ToChannelFromBotOAuthScope, conversationParameters, callback, cancellationToken);
        }
    }
}
