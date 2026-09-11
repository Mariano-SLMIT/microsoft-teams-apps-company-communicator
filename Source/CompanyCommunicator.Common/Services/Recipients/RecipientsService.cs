// <copyright file="RecipientsService.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Services.Recipients
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Threading.Tasks;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Extensions;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.SentNotificationData;
    using Microsoft.Teams.Apps.CompanyCommunicator.Common.Utilities;

    /// <summary>
    /// Recipients service.
    /// </summary>
    public class RecipientsService : IRecipientsService
    {
        private readonly ISentNotificationDataRepository sentNotificationDataRepository;

        /// <summary>
        /// Initializes a new instance of the <see cref="RecipientsService"/> class.
        /// </summary>
        /// <param name="sentNotificationDataRepository">sent notification data repository.</param>
        public RecipientsService(ISentNotificationDataRepository sentNotificationDataRepository)
        {
            this.sentNotificationDataRepository = sentNotificationDataRepository ?? throw new ArgumentNullException(nameof(sentNotificationDataRepository));
        }

        /// <inheritdoc/>
        public async Task<RecipientsInfo> BatchRecipients(IEnumerable<SentNotificationDataEntity> recipients)
        {
            if (recipients == null)
            {
                throw new ArgumentNullException(nameof(IEnumerable<SentNotificationDataEntity>));
            }

            var recipientsList = recipients.ToList();

            // Guard against an empty recipient list (e.g. group/team sync produced zero
            // valid recipients). Without this check, FirstOrDefault() returns null and
            // accessing .PartitionKey throws a NullReferenceException that aborts the
            // whole orchestration instead of reporting zero recipients cleanly.
            if (recipientsList.Count == 0)
            {
                return new RecipientsInfo(notificationId: null)
                {
                    TotalRecipientCount = 0,
                };
            }

            var notificationId = recipientsList.FirstOrDefault().PartitionKey;

            var recipientBatches = recipientsList.AsBatches(Constants.MaximumNumberOfRecipientsInBatch);
            var recipientInfo = new RecipientsInfo(notificationId)
            {
                TotalRecipientCount = recipientsList.Count,
            };
            int batchIndex = 1;
            foreach (var recipientBatch in recipientBatches)
            {
                var recipientBatchList = recipientBatch.ToList();

                // Update PartitionKey to Batch Key
                recipientBatchList.ForEach(s =>
                    {
                        s.PartitionKey = PartitionKeyUtility.CreateBatchPartitionKey(s.PartitionKey, batchIndex);

                        // Update if there is any recipient which has no conversation id.
                        recipientInfo.HasRecipientsPendingInstallation |= string.IsNullOrEmpty(s.ConversationId);
                    });

                // Store.
                await this.sentNotificationDataRepository.BatchInsertOrMergeAsync(recipientBatch);
                recipientInfo.BatchKeys.Add(recipientBatch.FirstOrDefault().PartitionKey);
                batchIndex++;
            }

            return recipientInfo;
        }
    }
}
