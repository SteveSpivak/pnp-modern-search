import * as React from "react";
import { ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClientFactory, MSGraphClientV3 } from "@microsoft/sp-http";

export interface ICopilotMeetingInsightsProps {
  meetingId: string;
  userId: string;
  serviceScope: ServiceScope;
}

export const CopilotMeetingInsights: React.FC<ICopilotMeetingInsightsProps> = ({ meetingId, userId, serviceScope }) => {
  const [insights, setInsights] = React.useState<any | null>(null);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>();
  const [graphClient, setGraphClient] = React.useState<MSGraphClientV3 | null>(null);

  React.useEffect(() => {
    let active = true;

    // Initialize MSGraphClient using the provided ServiceScope
    const initGraphClient = async () => {
      if (!serviceScope) return;

      try {
        const msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
        const client = await msGraphClientFactory.getClient("3");
        if (active) {
          setGraphClient(client);
        }
      } catch (err) {
        console.error("[CopilotMeetingInsights] Failed to initialize Graph Client.", err);
        if (active) {
          setError("Failed to initialize Microsoft Graph client.");
          setLoading(false);
        }
      }
    };

    initGraphClient();

    return () => { active = false; };
  }, [serviceScope]);

  React.useEffect(() => {
    let active = true;

    const fetchInsights = async () => {
      if (!meetingId || !userId) {
        setLoading(false);
        setError("Missing Meeting ID or User ID.");
        return;
      }

      if (!graphClient) {
        return;
      }

      try {
        // Utilizing v1.0 of the Meeting Insights API
        const endpoint = `/users/${userId}/onlineMeetings/${meetingId}/aiInsights`;
        const response = await graphClient.api(endpoint).version("v1.0").get();

        if (active) {
          setInsights(response);
          setLoading(false);
        }
      } catch (err: any) {
        console.error("[CopilotMeetingInsights] Error fetching meeting insights.", err);
        if (active) {
          setError(err.message || "Failed to retrieve meeting insights.");
          setLoading(false);
        }
      }
    };

    fetchInsights();

    return () => {
      active = false;
    };
  }, [meetingId, userId, graphClient]);

  return (
    <div style={{ padding: "16px", border: "1px solid #ccc", borderRadius: "8px", maxWidth: "400px", fontFamily: "Segoe UI, sans-serif" }}>
      <h3 style={{ marginTop: 0 }}>Copilot Meeting Insights</h3>
      {loading && <div>Loading insights...</div>}

      {error && <div style={{ color: "red" }}>Error: {error}</div>}

      {!loading && !error && insights && (
        <div>
          <h4>Summary:</h4>
          <p>{insights.summary || "No summary available."}</p>

          <h4>Action Items:</h4>
          {insights.actionItems && insights.actionItems.length > 0 ? (
            <ul>
              {insights.actionItems.map((item: any, idx: number) => (
                <li key={idx}>{item.task || item}</li>
              ))}
            </ul>
          ) : (
            <p>No action items found.</p>
          )}
        </div>
      )}

      {!loading && !error && !insights && <div>No meeting insights found.</div>}
    </div>
  );
};
