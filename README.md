# wyze-zendesk-ai-ticket-classifier

A Zendesk sidebar application built in partnership with Wyze that analyzes 
customer support call transcripts using GPT-4o and auto-populates ticket 
fields — reducing average tagging time from 3 minutes to under 30 seconds 
and projecting $302,400 in annual labor savings.

## What We Built

After a call ends, the agent clicks one button in the Zendesk sidebar. The 
app fetches the transcript, strips PII, and sends it to a Node.js backend 
running a multi-pass GPT-4o analysis engine. The engine performs a 
breadth-first traversal of the conditional ticket field hierarchy, scoping 
each prompt to only the fields unlocked at that stage, then writes results 
to the live Zendesk form incrementally after each pass. The agent reviews 
all suggested values with confidence scores before anything is committed — 
keeping a human in the loop on every decision.

## The Core Engineering Challenge

Sending the full taxonomy of 2,000+ tag values to GPT-4o in a single prompt 
caused the model to hallucinate non-existent field values. The multi-pass 
architecture solves this by breaking the analysis into conditional stages and 
presenting the model with only the options relevant at each level of the field 
hierarchy. A post-pass validation layer additionally rejects any returned value 
not present in the field's allowed option set.

## Results

- 94.8% field match rate against a gold-standard set of 20 tickets
- Tagging time reduced from 3 minutes to under 30 seconds per call
- Projected $302,400 in annual labor savings based on client-provided metrics

## Tech Stack

- **Frontend:** Zendesk Apps Framework (ZAF), vanilla JavaScript
- **Backend:** Node.js, Express
- **AI:** OpenAI GPT-4o
- **Data:** XLSX taxonomy loader, JSONL timing logs

## Setup

This repo contains the generalized source files. To run locally you will need:

1. A `.env` file with `OPENAI_API_KEY`, `ZENDESK_SUBDOMAIN`, `ZENDESK_EMAIL`, 
   and `ZENDESK_API_TOKEN`
2. A `data/taxonomy_data.xlsx` file with a `Field_options` sheet containing 
   a `Tag` column
3. `data/field-definitions.json`, `data/condition-lookup.json`, and 
   `data/form-config.json` with your Zendesk field configuration

Run `npm install` then `node server/server.js`.
