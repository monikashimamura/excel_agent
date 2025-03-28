from agent_network.base import BaseAgent
from agent_network.exceptions import ReportError
import agent_network.utils.storage.oss as oss
import requests


class ExcelReadAgent(BaseAgent):
    def __init__(self, graph, config, logger):
        super().__init__(graph, config, logger)

    def forward(self, messages, **kwargs):
        excel_file_url = kwargs.get("excel_file_url")
        if not excel_file_url:
            raise ReportError("excel_file_url is not provided", "AgentNetworkPlannerGroup/AgentNetworkPlanner")
        try:
            # 先尝试从OSS下载文件，如果失败则直接使用URL（也就是再尝试将url当作本地文件使用）。
            # 也可以多传一个参数，来控制是否从OSS下载文件。
            excel_file = oss.download_file(excel_file_url)
            use_oss = True
        except Exception as e:
            print(f"Failed to download file from OSS: {e}")
            excel_file = excel_file_url
            use_oss = False
        try:
            import openpyxl
            if not use_oss:
                workbook = openpyxl.load_workbook(excel_file)
            else:
                from tempfile import NamedTemporaryFile
                with NamedTemporaryFile() as tmp:
                    tmp.write(excel_file)
                    workbook = openpyxl.load_workbook(tmp)
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


class ExcelUploadAgent(BaseAgent):
    def __init__(self, graph, config, logger):
        super().__init__(graph, config, logger)

    def forward(self, messages, **kwargs):
        file_name = kwargs.get("excel_file_name")
        file_path = kwargs.get("excel_file_path")
        if not file_name:
            raise ReportError("excel_file_name is not provided", "AgentNetworkPlannerGroup/AgentNetworkPlanner")
        if not file_path:
            raise ReportError("excel_file_path is not provided", "AgentNetworkPlannerGroup/AgentNetworkPlanner")
        try:
            import oss2
            from oss2.credentials import StaticCredentialsProvider
            # 从环境变量中获取访问凭证
            auth = oss2.ProviderAuthV4(
                StaticCredentialsProvider
                ("LTAI5tDpwsApU5N4mi7iiRWn", "RypX56pi3HCXIgkpjLJHYxHTz19lzR"))

            # 设置Endpoint和Region
            endpoint = "https://oss-cn-hangzhou.aliyuncs.com"
            region = "cn-hangzhou"

            def generate_unique_bucket_name():
                return "agent"

            # 生成唯一的Bucket名称
            bucket_name = generate_unique_bucket_name()
            bucket = oss2.Bucket(auth, endpoint, bucket_name, region=region)

            with open(file_path, mode='rb') as f:
                content = f.read()
                bucket.put_object(file_name, content)
                url = bucket.sign_url('GET', file_name, 604800, slash_safe=True)
                print(url)
            result = {
                "excel_file_url": url
            }
            oss.upload_file(url, requests.get(url).content)
            return result
        except Exception as e:
            raise ReportError(f"ExcelUploadAgent error: {e}", "AgentNetworkPlannerGroup/AgentNetworkPlanner")
