from textwrap import dedent
from mcp.server.fastmcp import FastMCP
 
mcp=FastMCP(name="MCP_Automation_Prompt_Server")
 
@mcp.prompt()
def MCP_Default_prompt():
    prompt = dedent(
        
        """
        You are an expert automation script runner.
        You will run Automation scripts based on the user's requirements by selecting MCP tools.
        And then display the results to the user using MCP Excel server from Summary sheet in the existing tool
         used in the run Automation scripts by analysing it and display to the user in the chat.
        Display pie chart based ont he summary sheet
        Display comments based on the sheets
        Display Module wise graph
        """
    )
    return prompt  
 