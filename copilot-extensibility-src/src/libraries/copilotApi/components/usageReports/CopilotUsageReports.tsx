import * as React from "react";
import { ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClientFactory, MSGraphClientV3 } from "@microsoft/sp-http";

export interface ICopilotUsageReportsProps {
  period: "D7" | "D30" | "D90" | "D180";
  serviceScope: ServiceScope;
}

export const CopilotUsageReports: React.FC<ICopilotUsageReportsProps> = ({ period, serviceScope }) => {
  const [reports, setReports] = React.useState<any[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>();
  const [graphClient, setGraphClient] = React.useState<MSGraphClientV3 | null>(null);

  React.useEffect(() => {
    let active = true;

    const initGraphClient = async () => {
      if (!serviceScope) return;

      try {
        const msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
        const client = await msGraphClientFactory.getClient("3");
        if (active) {
          setGraphClient(client);
        }
      } catch (err) {
        console.error("[CopilotUsageReports] Failed to initialize Graph Client.", err);
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

    const fetchReports = async () => {
      if (!graphClient) {
        return;
      }

      setLoading(true);

      try {
        // The Accept header must be application/json to receive JSON instead of CSV
        const response = await graphClient
          .api(`/reports/getMicrosoft365CopilotUsageUserDetail(period='${period}')`)
          .version("beta")
          .header("Accept", "application/json")
          .get();

        if (active) {
          setReports(response.value || []);
          setLoading(false);
        }
      } catch (err: any) {
        console.error("[CopilotUsageReports] Error fetching usage reports.", err);
        if (active) {
          // If the user does not have Reports.Read.All permission, it will throw a 403.
          setError(err.message || "Failed to retrieve usage reports. Ensure you have 'Reports.Read.All' permission.");
          setLoading(false);
        }
      }
    };

    fetchReports();

    return () => {
      active = false;
    };
  }, [period, graphClient]);

  return (
    <div style={{ padding: "16px", border: "1px solid #ccc", borderRadius: "8px", maxWidth: "800px", fontFamily: "Segoe UI, sans-serif" }}>
      <h3 style={{ marginTop: 0 }}>Copilot Usage Reports ({period})</h3>
      {loading && <div>Loading usage reports...</div>}

      {error && <div style={{ color: "red", padding: "8px", backgroundColor: "#fde7e9", borderRadius: "4px" }}>Error: {error}</div>}

      {!loading && !error && reports.length > 0 && (
        <table style={{ width: "100%", borderCollapse: "collapse", textAlign: "left" }}>
          <thead>
            <tr style={{ borderBottom: "1px solid #ddd" }}>
              <th style={{ padding: "8px" }}>User Name</th>
              <th style={{ padding: "8px" }}>Total Copilot Actions</th>
              <th style={{ padding: "8px" }}>Last Activity Date</th>
            </tr>
          </thead>
          <tbody>
            {reports.map((report, idx) => (
              <tr key={idx} style={{ borderBottom: "1px solid #eee" }}>
                <td style={{ padding: "8px" }}>{report.userPrincipalName || report.displayName}</td>
                <td style={{ padding: "8px" }}>{report.totalCopilotActions || 0}</td>
                <td style={{ padding: "8px" }}>{report.lastActivityDate || "N/A"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {!loading && !error && reports.length === 0 && <div>No usage data found for the selected period.</div>}
    </div>
  );
};
