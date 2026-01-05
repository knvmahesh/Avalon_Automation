import asyncio
import streamlit as st
from langchain_mcp_adapters.client import MultiServerMCPClient
from langchain.chat_models import init_chat_model
from langchain_openai import ChatOpenAI
from langgraph.graph import StateGraph, MessagesState, START, END
from langchain_core.messages import SystemMessage
from langgraph.prebuilt import ToolNode
import os
from dotenv import load_dotenv
import openpyxl
from io import BytesIO
from datetime import datetime
#import excel_
import messagebox

import Excel
# Load .env file
load_dotenv()

#Getting Excel path from environment variable
path = os.getenv("EXCEL_FILE_PATH")

async def run_mcp_query(user_input):
    # Get keys from environment variables
    openai_key = os.getenv("OPENAI_API_KEY")
    #print(openai_key)
    # Model (using free tier OpenAI model here)
    model = ChatOpenAI(model="gpt-4o-mini", api_key=openai_key)
    print("Hi modal")
    #print(model)
    # MCP Client via HTTP
    client = MultiServerMCPClient(
        {
            "fusion_Automation": {
                "transport": "streamable_http",
                "url": "http://0.0.0.0:8007/mcp" # your MCP server URL
                 
            },
            "excel": {
                "transport": "streamable-http",
                "url": "http://0.0.0.0:8017/mcp"
      }
        }
        
    )
    
    print("Hi http")
    tools = await client.get_tools()
    print("Hi tools")
    model_with_tools = model.bind_tools(tools)
    print("Hi model_with_tools")
    sys_Prompt = """Task: Execute Fusion Automation and analyze the Excel file.

Workflow Constraints:

Single Execution: Trigger the Fusion Automation tool exactly once based on the user query. Do not initiate any retry loops, repetitive executions, or secondary tool calls if the first one fails.

Input Handling: The tool will provide the output as Excel file. Do not attempt to locate, open, or read external Excel files directly; rely solely on the text provided by the tool.

Data Validation: * If the tool output is empty, indicates an error, or states the sheet is not updated: Stop immediately. Inform the user: "No data is available for analysis."

If Excel file is present: Proceed to the analysis phase.

Analysis & Insights: Using Excel tool Analyze the 'Summary' sheet provided in the Excel file path returned from Fusion Atomation tool . Extract and summarize the key metrics into a clear, bulleted format within the chat.

Strict Guardrails: * No Fabrication: Use only the figures present in the Excel file.

Scope: Do not reference data from any sheet other than the 'Summary' sheet provided.

No Loops: After providing the insights or the "No data" message, terminate the process.

Goal: Provide a concise, high-level summary of automation results based strictly on the Excel file returned by the tool."""
    sys_msg = SystemMessage(content=sys_Prompt)
    tool_node = ToolNode(tools)
    #print(tool_node)
    #messagebox.showinfo("Success", tool_node)
    def should_continue(state: MessagesState):
        messages = state["messages"]
        last_message = messages[-1]
        print("Last message in should_continue:", last_message)    
        if last_message.tool_calls:
            print(last_message.tool_calls)
            return "tools"
        return END

    async def call_model(state: MessagesState):
        messages = state["messages"]+[sys_msg]
        response = await model_with_tools.ainvoke(messages)
        print("Response in call_model:", response)
        return {"messages": [response]}

    # LangGraph pipeline
    builder = StateGraph(MessagesState)
    builder.add_node("call_model", call_model)
    builder.add_node("tools", tool_node)
    builder.add_edge(START, "call_model")
    builder.add_conditional_edges("call_model", should_continue)
    builder.add_edge("tools", "call_model")

    graph = builder.compile()
    result = await graph.ainvoke({"messages": [{"role": "user", "content": user_input}]},config={"recursion_limit": 5})

    # Extract last message text
    last_msg = result["messages"][-1].content
    return last_msg if isinstance(last_msg, str) else str(last_msg)


def main():
    st.set_page_config(page_title="Fusion Automation Chat", page_icon="Avalon Logo.ico")
    st.title("Avalon Tool(Streamlit)")
    st.info("Type input in chat to run the script like for Eg:'Run HCM Employee Creation Automation Script '")
    user_input = st.text_input("Run Automation script")
    if st.button("Send") and user_input.strip():
        with st.spinner("Running..."):
            answer = asyncio.run(run_mcp_query(user_input))
            st.success(answer)
        
        wb = openpyxl.load_workbook(path)
        excel_bytes = Excel.workbook_to_bytes(wb)
        st.download_button(
                 label="Download Excel File",
                 data = excel_bytes,
                file_name=f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            

if __name__ == "__main__":
    main()

