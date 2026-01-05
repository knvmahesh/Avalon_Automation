from mcp.server.fastmcp import FastMCP
from dotenv import load_dotenv
import openpyxl
import HCM_Automation
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException,NoSuchElementException

load_dotenv("../.env")

# Create an MCP server
mcp = FastMCP(name="Test_Automation",host="0.0.0.0",port=8050)

#mcp = FastMCP(name="Test_Automation",port=8000)

print("Server Module Loaded")   
@mcp.tool()
def RunAutomationScript():
    """Run HCM Employee Creation Automation Script"""
    HCM_Automation.Manage_Jobs()
    HCM_Automation.Manage_Departments()
    HCM_Automation.Manage_Positions()
    HCM_Automation.Employee_Creation()  
    HCM_Automation.Termination_Employee()

if __name__ == "__main__":
    transport = "stdio"
    if transport == "stdio":
        print("Running server with stdio transport")
        mcp.run(transport="stdio")
    elif transport == "sse":
        print("Running server with SSE transport")
        mcp.run(transport="sse")
    elif transport == "streamable-http":
        print("Running server with Streamable HTTP transport")
        mcp.run(transport="streamable-http")
    else:
        raise ValueError(f"Unknown transport: {transport}")
