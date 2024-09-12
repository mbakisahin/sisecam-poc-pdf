from utils.system_messages import SYSTEM_MESSAGE_COMPARISON, SYSTEM_MESSAGE_COMPARISON_CHUNK
from utils.openAI import OpenAIClient
from utils.send_email import EmailClient

from utils.excel_generator import ExcelReportGenerator


class SummaryComparator:
    """
    A class to handle the comparison of summaries between an original document and its nearest neighbors using OpenAI.
    It also generates Excel reports and sends them via email.
    """

    def __init__(self, engine="gpt-4o"):
        """
        Initializes the SummaryComparator with a specified OpenAI engine. It also initializes instances for email
        sending and Excel report generation.

        :param engine: The OpenAI engine to use for the comparisons (default: gpt-4o).
        """
        self.engine = engine
        self.openai_client = OpenAIClient(self.engine)
        self.email_client = EmailClient()
        self.excel_report_generator = ExcelReportGenerator()  # Excel operations using a new class

    def compare_with_multiple_neighbors(self, original_file_name, original_summary, neighbors, metadata=None):
        """
        Compares the original document's summary with multiple neighbor documents' summaries.
        It also generates an Excel report and sends the comparison via email.

        :param original_file_name: The name of the original file being compared.
        :param original_summary: The summary of the original document.
        :param neighbors: A list of dictionaries containing neighbor document summaries and URLs.
        :param metadata: Optional metadata related to the comparison (e.g., keyword, URL, date).
        """
        # Combine the summaries of all neighbor documents into a single string
        combined_neighbors_summary = "\n\n".join([neighbor['summary'] for neighbor in neighbors])
        neighbor_urls = [neighbor.get('url', '#') for neighbor in neighbors]

        # Perform a combined comparison between the original summary and all neighbors
        combined_comparison = self.compare_summaries(
            original_summary=original_summary,
            neighbor_summary=combined_neighbors_summary,
            system_messages=SYSTEM_MESSAGE_COMPARISON,
            accumulate=True
        )

        # Perform individual comparisons between the original summary and each neighbor summary
        individual_comparisons = []
        for neighbor in neighbors:
            comparison_result = self.compare_summaries(
                original_summary=original_summary,
                neighbor_summary=neighbor['summary'],
                system_messages=SYSTEM_MESSAGE_COMPARISON_CHUNK,
                accumulate=True
            )
            individual_comparisons.append(comparison_result)

        # Prepare the metadata dictionary for the Excel report
        metadata_dict = {
            'combined_comparison': combined_comparison,
            'individual_comparisons': individual_comparisons,
            'keyword': metadata.get('keyword', 'N/A') if metadata else 'N/A',
            'url': metadata.get('URL', 'N/A') if metadata else 'N/A',
            'date': metadata.get('notified_date', 'N/A') if metadata else 'N/A',
            'neighbor_urls': neighbor_urls
        }

        # Use the ExcelReportGenerator class to create the Excel report
        excel_file_path = 'comparison_report.xlsx'
        self.excel_report_generator.create_excel(metadata_dict, excel_file_path)

        # Send the email with the report as an attachment
        subject = f"Summary Comparison Results for {original_file_name} vs Neighbors"
        body = "Please find attached the comparison report in Excel format."
        self.email_client.send_email(subject, body, 'recipient@example.com', excel_file_path)

    def compare_summaries(self, original_summary, neighbor_summary, system_messages, accumulate=False):
        """
        Compares the original summary with a single neighbor summary using OpenAI in a single comparison.

        :param original_summary: The summary text of the original document.
        :param neighbor_summary: The summary text of the neighbor document.
        :param system_messages: System messages that guide the comparison logic.
        :param accumulate: If True, accumulates the comparison result for aggregation (default: False).
        :return: The combined differences between the original summary and the neighbor summary.
        """
        # Prepare the input text for comparison
        input_text = (
            f"Original Summary:\n{original_summary}\n\n"
            f"Neighbor Summary:\n{neighbor_summary}\n\n"
            f"Please provide the key differences between the original summary and the neighbor summary."
        )

        # Perform the comparison using OpenAI
        comparison_result = self.openai_client.compare_texts(input_text, system_messages)

        # If accumulate is True, return the raw comparison result
        if accumulate:
            return comparison_result
