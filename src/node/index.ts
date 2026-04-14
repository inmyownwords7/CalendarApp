import { WebClient } from "@slack/web-api";
import { Client as NotionClient } from "@notionhq/client";

console.log("Node dev entrypoint running.");

const slack = new WebClient(process.env.SLACK_BOT_TOKEN);
const notion = new NotionClient({ auth: process.env.NOTION_TOKEN });

// Put your Node-side experimentation here.
