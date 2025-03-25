from agent_network.base import BaseAgent


class Agent1(BaseAgent):
    def __init__(self, graph, config, logger):
        super().__init__(graph, config, logger)

    def forward(self, message, **kwargs):
        messages = []
        self.add_message("user", f"task: {kwargs['task']}", messages)
        response = self.chat_llm(messages)
        print('response: ' + response.content)
        results = {
            "result": response.content,
        }
        return results
