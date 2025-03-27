from agent_network.base import BaseAgent
from agent_network.exceptions import ReportError
import agent_network.graph.context as ctx


class ExcelReadAgent(BaseAgent):
    def __init__(self, graph, config, logger):
        super().__init__(graph, config, logger)

    def forward(self, messages, **kwargs):
        # excel_file_name = kwargs.get("excel_file_name")
        excel_file_name = ctx.retrieves_all()['excel_file_path']
        try:
            if not excel_file_name:
                raise ReportError("excel_file_name is not provided", "AgentNetworkPlannerGroup/AgentNetworkPlanner")

            import openpyxl
            workbook = openpyxl.load_workbook(excel_file_name)
            sheet1 = workbook['Sheet1']
            read = True if sheet1['A1'].value else False
            recognize = not read
            all_value = []
            for row in sheet1.rows:
                row_value = [cell.value if read else cell.fill.fgColor.value for cell in row]
                all_value.append(row_value)
            if recognize:
                result = all_value
            else:
                key = all_value[0]
                all_value.pop(0)
                value = all_value
                print(key, value)
                result = [dict(zip(key, row)) for row in value]
            excel_read_result = str(result)
            print(excel_read_result)
            self.log("assistant", excel_read_result)

            result = {
                "excel_read_result": excel_read_result
            }
            return result
        except Exception as e:
            raise ReportError(f"ExcelReadAgent error: {e}", "AgentNetworkPlannerGroup/AgentNetworkPlanner")
