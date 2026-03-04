import * as React from "react";
import { CopilotApiLibrary } from "../../CopilotApiLibrary";

export interface ICopilotMeetingInsightsProps {
  meetingId: string;
  userId: string;
}

export const CopilotMeetingInsights: React.FC<ICopilotMeetingInsightsProps> = ({ meetingId, userId }) => {
  const [insights, setInsights] = React.useState<any | null>(null);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>();

  React.useEffect(() => {
    let active = true;

    const fetchInsights = async () => {
      if (!meetingId || !userId) {
        setLoading(false);
        setError("Missing Meeting ID or User ID.");
        return;
      }

      try {
        const client = CopilotApiLibrary.msGraphClient;
        if (!client) {
          throw new Error("Microsoft Graph client is not initialized.");
        }

        // Utilizing v1.0 of the Meeting Insights API
        const endpoint = `/users/${userId}/onlineMeetings/${meetingId}/aiInsights`;
        const response = await client.api(endpoint).version("v1.0").get();

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
  }, [meetingId, userId]);

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
