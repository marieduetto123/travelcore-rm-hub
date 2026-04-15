# TravelCore RM Hub - User Testing Plan
## Calendar: Monthly & Weekly Views

**Version:** v37 | **Date:** April 15, 2026
**Test URL:** https://travelcore-rm-hub.vercel.app

---

## Core Purpose

The calendar exists to give Revenue Managers all the data they need — occupancy, rates, availability, meal plan mix, operator allotments, events, comparisons — to make confident close-out decisions. Every feature (heatmaps, filters, compare modes, cell metrics, eye popup, weekly grid) serves one goal: **surface the right information so the user can decide what to close, when, and for whom.**

This testing plan evaluates whether the calendar achieves that goal — whether users can move from data to close-out action quickly, accurately, and with confidence.

---

## Test Objectives

1. Validate that the calendar surfaces the right data to support close-out decisions
2. Measure how quickly users can move from "seeing a problem" to "executing a close-out"
3. Identify gaps where users need information the calendar doesn't provide (or doesn't surface clearly enough)
4. Test whether the data hierarchy (monthly overview → day detail → weekly grid) matches how RMs actually think about close-outs
5. Assess whether heatmaps, filters, and compare modes genuinely accelerate close-out decision-making or add friction

---

## Participant Profile

- **Role:** Revenue Manager, Cluster RM, or Reservations Manager
- **Experience:** 1+ year using a PMS or RMS; regularly performs stop-sell / close-out actions
- **Sample size:** 5-8 participants
- **Session length:** 45-60 minutes

---

## Pre-Test Setup

- Browser: Chrome (latest), 1440px+ viewport
- Start on the Calendar tab, Monthly view, 3-month mode
- Compare mode: vs STLY (default)
- All filters cleared
- No heatmap active

---

## Section 1: Scanning for Close-Out Signals (5 min)

### Scenario 1.1 - Morning Close-Out Check
> *"It's 8am Monday morning. Your first job is always to scan the next 2-3 weeks and decide if anything needs to be closed out today. Walk me through what you're looking at and what would trigger a close-out decision."*

**Observe:**
- [ ] Does the user orient to occupancy levels as the primary close-out signal?
- [ ] Do they notice which days are already locked?
- [ ] Do they spot days with action-needed indicators (high hotel, low TO)?
- [ ] Do they mention rates, remaining inventory, or events as factors?

**Follow-up:**
- "You've been scanning for 30 seconds. Your GM walks over and asks 'Do we need to close anything today?' Based on what you see, what would you tell them and why?"
- "What specific data point on this screen would make you say 'That day needs to be closed out right now'?"
- "Is there anything missing from this view that you'd normally check before making a close-out call?"

---

## Section 2: Building the Close-Out Picture - Navigation (10 min)

### Scenario 2.1 - Peak Season Close-Out Planning
> *"Summer season (June-August) is approaching and your Director has asked you to review the calendar and flag any dates that might need preemptive close-outs. You need to scan all 3 months and identify the pressure points. How would you get to that view?"*

**Observe:**
- [ ] Do they use the date range picker, arrows, or month-view dropdown?
- [ ] Can they operate the date range picker without guidance?
- [ ] Do they discover the quick-select presets?

**Follow-up:**
- "You're now looking at June-August. You notice several days above 85% occupancy. Before deciding to close them out, what other information would you want to see alongside occupancy?"
- "Your Director says 'Give me the whole second half of the year — I want to see where we'll have close-out pressure.' How would you get that overview?"
- "In the 6-month view, you can see colour coding but no numbers. Is that enough to identify which days need close-out attention, or do you need more detail?"

### Scenario 2.2 - Configuring Metrics for Close-Out Decisions
> *"You're preparing for a close-out review meeting. You want each day cell to show you the 3-4 data points that matter most for deciding whether to close. What metrics would you choose, and how would you set that up?"*

**Observe:**
- [ ] Do they find the Cell Metrics button?
- [ ] Which metrics do they prioritise for close-out decisions?
- [ ] Do they understand the Hotel vs T distinction and why it matters for close-outs?

**Follow-up:**
- "You've chosen your metrics. Walk me through a single day cell now — based on these numbers, would you close this day? What's your reasoning?"
- "You want to see whether a specific operator is over their allotment — that might justify closing just their channel. Can you configure the view to show individual operator data?"

---

## Section 3: Gathering Close-Out Intelligence (10 min)

### Scenario 3.1 - Comparing Against Forecast to Justify Close-Outs
> *"Your GM has a rule: 'Don't close anything unless we're tracking ahead of forecast by at least 10%.' You need to identify which days are significantly ahead of forecast and might justify a close-out. How would you set up the view to answer that?"*

**Observe:**
- [ ] Can they switch to Forecast compare mode?
- [ ] Do they understand the green/red colouring as above/below forecast?
- [ ] Can they identify days that are ahead of forecast?

**Follow-up:**
- "You see a day showing '92% / 82%' — 10 points ahead of forecast. That meets the GM's threshold. But before you close it out, what else on this screen would you check?"
- "Another day shows '65% / 70%' in red — behind forecast. Would you ever close out a day that's behind forecast? Under what circumstances?"
- "Your GM now says 'How did we compare to the same time last year?' Switching to that comparison, does it change your close-out recommendation for any of these days?"

### Scenario 3.2 - Quick Day Assessment Before Close-Out
> *"You've identified a day that looks like it needs a close-out. Before you pull the trigger, you want to check the full breakdown — room types, rates, meal plans, operator allotments — to decide exactly what to close and for whom. How would you get that detail without leaving the monthly view?"*

**Observe:**
- [ ] Do they discover the eye icon on hover?
- [ ] Do they use it to check availability, rates, and close-out status?
- [ ] Can they identify which room types or operators to target for the close-out?

**Follow-up:**
- "Looking at the popup, you see Suites are nearly sold out but Standard rooms still have availability. Would you close all room types or just Suites? What data here informed that decision?"
- "You notice one operator's allotment is almost fully used while another has barely touched theirs. Would that change who you close out? How?"
- "You've made your decision. Now you want to see this day in context with the surrounding days before executing. How would you get to the weekly view?"

### Scenario 3.3 - Events Driving Close-Out Timing
> *"Your Sales team has told you a major conference is happening next week. Events like this typically justify preemptive close-outs to protect rate integrity. How would you verify the event is in the system, and how would it influence your close-out decisions for those dates?"*

**Observe:**
- [ ] Can they find the event icon and tooltip?
- [ ] Do they use the event information alongside occupancy data to justify close-outs?

**Follow-up:**
- "You've confirmed the event is on March 12-14. Occupancy is already at 78% for those days. Would you close out now or wait? What additional data on this screen would inform that timing decision?"
- "The event tooltip shows it's a 'One-time' event. If it were 'Recurring', would that change your close-out approach?"

### Scenario 3.4 - Capacity Gut-Check
> *"A tour operator has requested 40 rooms for next Tuesday. You need to quickly check remaining capacity before deciding whether to accept the booking or close out their channel. How would you check without opening the full detail view?"*

**Observe:**
- [ ] Do they hover to trigger the capacity tooltip?
- [ ] Can they read remaining rooms?

**Follow-up:**
- "The tooltip shows 38 rooms remaining. That's tight for a 40-room request. What's your next step — accept, negotiate, or close out? What would you check next before deciding?"

---

## Section 4: Heatmaps for Close-Out Targeting (8 min)

### Scenario 4.1 - Visual Close-Out Audit
> *"Your Director wants a quick visual answer to: 'Show me the close-out landscape for the next 3 months — where are we open, where are we partially closed, and where are we fully locked?' Set that up."*

**Observe:**
- [ ] Do they find the Heatmap button and select Stop Sales?
- [ ] Do they understand the red/yellow/green colour coding?
- [ ] Can they use the heatmap to quickly identify patterns (clusters of closures, gaps)?

**Follow-up:**
- "You notice a cluster of red (fully closed) days in mid-April. Your Director asks 'Are the Suites specifically closed on those days, or is it all room types?' How would you drill into Suite-specific closures?"
- "Based on this heatmap, your Director says 'That Wednesday in the middle of all those closures is still green — should we close it too?' How would you decide?"

### Scenario 4.2 - Using Occupancy Heatmap to Identify Close-Out Candidates
> *"Instead of looking at what's already closed, you want to find days that should be closed but aren't yet. You want to visually spot all days running above 90% occupancy — those are your close-out candidates. How would you set that up?"*

**Observe:**
- [ ] Can they switch to Hotel Occupancy heatmap?
- [ ] Can they set the threshold to 90%?
- [ ] Do they use the results to identify close-out targets?

**Follow-up:**
- "You've found 8 days above 90%. Your GM says 'But only close the ones where we have fewer than 10 rooms left.' Can you add that as a condition to narrow the list?"
- "Now you have 4 days highlighted. Walk me through how you'd go from this heatmap view to actually executing the close-outs for those specific days."

---

## Section 5: Executing Close-Outs from Monthly View (8 min)

### Scenario 5.1 - Assessing Existing Close-Outs
> *"You've just taken over this property from a colleague. Before making any new close-out decisions, you need to understand what's already in place. How would you get a picture of the current close-out state?"*

**Observe:**
- [ ] Do they notice lock icons and the close-out count badge?
- [ ] Do they check individual locked days for details (room types, operators)?
- [ ] Do they use the heatmap or the eye popup to audit existing rules?

**Follow-up:**
- "Your colleague closed out 48 dates. Your GM says 'That seems aggressive — review them and reopen anything that doesn't need to be closed.' How would you approach that review? What data would you check for each closed day?"

### Scenario 5.2 - Range Close-Out
> *"You've reviewed the data and decided: Suite and Jr. Suite rooms need to be closed for March 20-25 because a tour operator has overbooked their allocation. Execute that close-out now."*

**Observe:**
- [ ] Do they use Select Range or Custom workflow?
- [ ] Can they click start/end dates on the grid?
- [ ] Can they configure room types and operator in the modal?

**Follow-up:**
- "The operator calls back — it's actually only March 20-22 that need closing. How would you adjust?"
- "Your GM says 'Close those room types, but only for the Sunshine Tours channel — keep them open for everyone else.' Can you set up that targeted rule?"
- "You've executed the close-out. How would you verify it applied correctly? What would you check?"

### Scenario 5.3 - Reopening After Close-Out
> *"The overbooking issue is resolved. You need to reopen March 20-22 but keep March 23-25 closed. How would you selectively reopen just those 3 days?"*

**Observe:**
- [ ] Do they find the Bulk Reopen button?
- [ ] Can they select only the intended days?
- [ ] Do they verify the result?

**Follow-up:**
- "How confident are you that you only reopened the correct 3 days? Walk me through how you'd double-check."
- "If you accidentally reopened March 23 as well, how would you fix it?"

---

## Section 6: Monthly Summary Supporting Close-Out Strategy (5 min)

### Scenario 6.1 - Close-Out Strategy Meeting
> *"You're presenting your close-out strategy for the next quarter to your Revenue Director. You need to back up your decisions with data — occupancy trends, rate performance, operator mix, and allotment usage. Where would you find the summary data to support your case?"*

**Observe:**
- [ ] Do they scroll to the monthly summary accordion?
- [ ] Can they find occupancy trends, ADR, revenue, and operator rates?
- [ ] Do they use the period average for cross-month comparison?

**Follow-up:**
- "Your Director asks 'What's the occupancy trend across the quarter — are we trending up or down?' Where would you find that?"
- "They then ask 'Which operators are contributing the most? Should we close any specific channels?' Can you answer that from the summary?"
- "Based on this summary, what close-out recommendations would you make for next quarter?"

---

## Section 7: Drilling Into the Weekly View for Close-Out Detail (3 min)

### Scenario 7.1 - Moving from Signal to Action
> *"On the monthly view, you've spotted that March 12th is at 91% occupancy with a conference event. You want to see the full week to decide if you should close just that day or a broader range. How would you drill in?"*

**Observe:**
- [ ] Do they click the cell to transition to weekly?
- [ ] Does the weekly view give them the context they need for the close-out decision?
- [ ] Do they notice the surrounding days' data?

**Follow-up:**
- "You're now in the weekly view. March 12th looks like it needs closing, but March 11th and 13th are also above 85%. Would you close all three or just the one? What data on this screen is guiding that decision?"
- "You've decided. How would you get back to the monthly view to check if there are other dates that need attention?"

---

## Section 8: Weekly View - Deep Close-Out Analysis (10 min)

### Scenario 8.1 - Reading the Full Close-Out Picture
> *"You're in the weekly view looking at 7 days of data. Your job right now is to decide: which of these 7 days need close-outs, what exactly should be closed on each day, and why. Walk me through your thought process as you read this grid."*

**Observe:**
- [ ] Do they start with the Close Outs group to see what's already in place?
- [ ] Do they cross-reference occupancy, availability, and rates before deciding?
- [ ] Do they consider operator-specific and room-type-specific close-outs?
- [ ] Do they use the row hierarchy (group → section → sub-row) effectively?

**Follow-up:**
- "You've identified 2 days that need close-outs. What specific data rows did you look at to reach that conclusion?"
- "There's a lot of data here. If you could only see 3 data groups to make close-out decisions, which would you keep?"

### Scenario 8.2 - Segment-Level Close-Out Decision
> *"Thursday's occupancy is 88%, but when you dig into the sub-rows, you see TO bookings are at 45% while Other Segments are at 43%. The TO allotment for one operator is almost fully used. Walk me through whether you'd close anything, and if so, what specifically."*

**Observe:**
- [ ] Do they expand Occupancy sub-rows to see segment breakdown?
- [ ] Do they check Room Availability for room-type detail?
- [ ] Do they check Travel Co. Rates for allotment usage?
- [ ] Do they consider closing a specific operator vs. a specific room type vs. the whole day?

**Follow-up:**
- "You've decided to close Suites for Sunshine Tours only. But the trend badge shows occupancy is '-12%' versus last year. Does that change your decision? How do you weigh current demand against historical comparison?"
- "What if the forecast shows Thursday should be at 95%? Would you close more aggressively or hold off?"

### Scenario 8.3 - Using Forecast Data to Time Close-Outs
> *"Your Finance team wants you to close out any day that's tracking 10+ points ahead of forecast — that's their threshold for 'close and protect the rate'. Set up the view to identify those days, then execute."*

**Observe:**
- [ ] Do they switch to Forecast compare mode?
- [ ] Can they identify days above forecast by 10+ points using trend badges?
- [ ] Do they execute the close-out from the weekly view?

**Follow-up:**
- "You see Monday is 8% ahead of forecast — close to the threshold but not quite. Would you preemptively close it? What other data would tip you over?"
- "In the sub-rows, you see small 'Fc:' forecast values. Do they help you make a more granular close-out decision (e.g., close one room type but not another)?"

---

## Section 9: Weekly Close-Out Execution (5 min)

### Scenario 9.1 - Quick Close-Out Status Check
> *"You're about to join a 5-minute standup. You need a fast answer: 'What's the close-out status for each of the next 7 days?' You don't need details — just the headline. How would you get that in under 10 seconds?"*

**Observe:**
- [ ] Do they collapse the Close Outs group to see the summary?
- [ ] Can they read the per-day labels (Full Close Out / 3 rules / Open)?
- [ ] Is the collapsed view sufficient for a quick status check?

**Follow-up:**
- "Wednesday shows '3 rules'. After the standup, your GM asks 'Which room types are affected by those 3 rules?' How would you find out?"
- "Is the collapsed summary the right amount of information for a quick check? Would you change what it shows?"

### Scenario 9.2 - Batch Close-Out from Weekly Grid
> *"After your analysis, you've decided Friday, Saturday, and Sunday all need to be closed for Beach Hols operator. Select those 3 days and execute the close-out."*

**Observe:**
- [ ] Do they click column headers to select days?
- [ ] Do they notice the lock icons changing colour?
- [ ] Does the "Close Out N Days" button appear?
- [ ] Can they configure operator-specific rules in the modal?

**Follow-up:**
- "You accidentally selected Thursday too. How would you deselect just that one day without starting over?"
- "Saturday's lock icon was already amber before you clicked. What does that tell you about the existing close-out state for that day? Would it affect your decision?"

---

## Section 10: Room-Level Close-Out Decisions (5 min)

### Scenario 10.1 - Room Type Close-Out Evaluation
> *"A tour operator wants to book 15 Deluxe rooms for Wednesday. You need to decide: accept the booking, or close out Deluxe rooms to protect remaining inventory. Check the data and make your call."*

**Observe:**
- [ ] Do they navigate to Room Availability > Deluxe?
- [ ] Do they check: available rooms, tentative sold, out-of-order, allotment remaining?
- [ ] Do they cross-reference with Travel Co. Rates for the operator's rate and promo?
- [ ] Do they make a clear close-out recommendation based on the data?

**Follow-up:**
- "Deluxe shows '8 avail / 27' but also 3 rooms as 'Tentative Sold (Group)'. If that tentative group confirms, you'd only have 5 real rooms left. Would you close out now or wait? What's the risk either way?"
- "The bar chart has 5 coloured segments. Which segments matter most for your close-out decision?"
- "The operator's rate shows $185 with EBB 10%. If you close Deluxe for this operator, you lose that revenue. How do you weigh close-out protection against the revenue loss?"

### Scenario 10.2 - Allotment-Driven Close-Out
> *"It's end of month and you're reviewing operator allotments. If an operator has nearly used their full allotment, you might need to close their channel to prevent overbooking. Where would you check, and what would trigger a close-out?"*

**Observe:**
- [ ] Do they navigate to Travel Co. Rates?
- [ ] Can they interpret allotment remaining figures?
- [ ] Do they identify operators needing channel closure?

**Follow-up:**
- "Beach Hols has only 2 rooms remaining in their allotment for Tuesday. Would you close their channel for that day? What if they have a high-value group enquiry coming in?"

---

## Section 11: Configuring the View for Close-Out Workflow (5 min)

### Scenario 11.1 - Building Your Close-Out Dashboard
> *"You do close-out reviews every morning. You want the weekly view set up so the most important close-out data is at the top and anything irrelevant is hidden. Configure the view for your ideal close-out workflow."*

**Observe:**
- [ ] Do they find Table Settings?
- [ ] What order do they put groups in for close-out work?
- [ ] What do they choose to hide?

**Follow-up:**
- "What order did you choose and why? Walk me through your close-out thinking process — what do you check first, second, third?"
- "If you close and reopen this tomorrow, would you expect your close-out layout to be saved?"

### Scenario 11.2 - Operator-Specific Close-Out Investigation
> *"Sunshine Tours has raised a concern that their All Inclusive Suite allocation is being closed too aggressively. You need to look at the data filtered to just their bookings, Suite rooms, and AI meal plan to decide if the current close-outs are justified. Set that up."*

**Observe:**
- [ ] Can they apply all three filters?
- [ ] Do they check close-out status, occupancy, and allotment for the filtered view?
- [ ] Can they determine whether current close-outs are justified?

**Follow-up:**
- "Based on the filtered data, are the current close-outs justified for Sunshine Tours, or should some be reopened?"
- "You've finished the investigation. Now clear filters and check if the close-out picture looks different for the full hotel. Does removing the filter change your recommendation?"

---

## Section 12: Close-Out Heatmap Matrix (3 min)

### Scenario 12.1 - Room Type Closure Audit
> *"Your Revenue Director wants a single view that answers: 'For each room type, which days this week are closed and which are open?' They want to spot gaps in the close-out strategy. Find that view."*

**Observe:**
- [ ] Can they switch to the Close Outs sub-tab in weekly view?
- [ ] Do they understand the room-type × day matrix?
- [ ] Can they spot close-out gaps (e.g., Suites closed Monday-Wednesday but open Thursday)?

**Follow-up:**
- "The Director points at a gap — Suites are closed Monday-Wednesday but open Thursday, then closed again Friday. Is that intentional or a mistake? How would you decide?"
- "Based on this matrix, what close-out actions would you recommend?"

---

## Section 13: End-to-End Close-Out Workflow (5 min)

### Scenario 13.1 - Competitive Pressure Close-Out
> *"It's Tuesday morning. Your competitor has just dropped rates by 15%. You need to protect your rate integrity by closing the right inventory. Starting from the monthly view: (1) identify which days this week are most vulnerable (high occupancy + low remaining inventory), (2) decide what to close for each day, and (3) execute the close-outs. Go."*

**Observe the full flow:**
- [ ] Do they scan monthly for the high-pressure days?
- [ ] Do they drill into the weekly view for detail?
- [ ] Do they cross-reference occupancy, availability, and rates before deciding?
- [ ] Can they execute targeted close-outs (specific room types/operators)?
- [ ] How many clicks from "I see a problem" to "close-out confirmed"?
- [ ] Any backtracking, confusion, or missing data?

**Follow-up:**
- "That took you [X] steps. In your current system, how many steps and screens would the same close-out workflow require?"
- "Was there a moment where you needed data that wasn't visible and had to go looking for it?"
- "If you could change one thing about the close-out workflow, what would it be?"

### Scenario 13.2 - Close-Out Handover
> *"You're going on holiday for a week. Your cover needs to manage close-outs while you're away. Using this tool, brief them on: what's currently closed, what's likely to need closing this week, and how to execute a close-out if needed."*

**Observe:**
- [ ] What data do they use to brief the handover?
- [ ] Do they naturally gravitate to the close-out status views?
- [ ] Can they clearly explain the close-out workflow to a hypothetical colleague?

**Follow-up:**
- "If your cover had never used this tool, could they figure out how to close something out? Where would they struggle?"
- "Is there anything you'd want to leave as a note or flag on specific days — like 'Watch this day, may need closing if occupancy hits 90%'?"

---

## Post-Test Close-Out Focused Questions

### Close-Out Confidence
1. "Thinking about every close-out decision you made during this session — how confident were you that you had enough data to make the right call each time?"
2. "Was there a moment where you hesitated before closing something out? What additional information would have made you more confident?"
3. "On a scale of 1-5, how much faster would this tool make your daily close-out review compared to your current process?"

### Data Completeness for Close-Outs
4. "Think about the last close-out you made in real life. Was there a piece of information you used that this tool doesn't show?"
5. "Which single data point on this calendar most often triggers your close-out decisions in practice?"
6. "The calendar shows occupancy, rates, availability, meal plans, operator allotments, events, and historical comparisons. Is anything missing that you'd need for a close-out decision?"
7. "You saw 'Tentative Sold (Group)' in room availability. How important is that for close-out timing — would a tentative group hold make you close sooner?"

### Close-Out Workflow Efficiency
8. "Think about the fastest close-out you executed during this session. Could it have been faster? How?"
9. "You used both the monthly view and weekly view for close-outs. If you could only use one, which would you choose for close-out decisions and why?"
10. "The tool lets you close by date range, by room type, by operator, and by meal plan. In practice, which of those dimensions do you close on most often?"
11. "Is the bulk reopen workflow adequate? What would make reopening closed dates easier?"

### Strategic Close-Out Support
12. "Your GM asks 'Show me one screen that tells me if we need to close anything today.' Which view and setup would you show them?"
13. "If this tool could send you a morning alert saying 'These 3 days need close-out attention', what criteria should it use?"
14. "You used the heatmap to visualise close-out status. Would you use that daily, or is the standard calendar view enough?"
15. "Is there a close-out scenario you deal with regularly that this tool doesn't handle well?"

### Integration & Gaps
16. "If this tool could connect to one other system to improve your close-out workflow (PMS, channel manager, operator portal, etc.), which integration would save you the most time?"
17. "After executing a close-out here, what's the next step in your real workflow? Does this tool cover that, or is there a gap?"
18. "Is there any close-out data you'd want to export or share? Who would receive it and in what format?"

---

## Scoring Rubric

| Metric | Target | Measurement |
|--------|--------|-------------|
| Time from signal to close-out | < 90 seconds | From spotting a problem day to confirmed close-out |
| Close-out accuracy | > 90% | Users close the correct room types / operators / dates |
| Data sufficiency rating | > 4/5 | "Did you have enough data to make a confident close-out decision?" |
| Close-out task completion | > 85% | Close-out tasks completed without facilitator help |
| Heatmap-to-close-out flow | > 70% | Users who can go from heatmap insight to close-out execution unaided |
| Targeted close-out success | > 75% | Users who successfully close specific room types or operators (not just full day) |
| Reopen accuracy | > 85% | Users who reopen only intended dates without affecting others |
| Eye popup utility for close-outs | Tracked | Whether users check the popup before executing a close-out |
| Preferred close-out view | Tracked | Monthly vs. weekly — which view users prefer for close-out work |
| SUS score (System Usability Scale) | > 75 | Post-test SUS questionnaire |

---

## Notes Template

| Participant | Date | Role | Close-Out Approach | Data Gaps Mentioned | Workflow Friction | Feature Requests |
|-------------|------|------|--------------------|--------------------|--------------------|------------------|
| P1 | | | | | | |
| P2 | | | | | | |
| P3 | | | | | | |
| P4 | | | | | | |
| P5 | | | | | | |

### Close-Out Workflow Timing Log

| Scenario | P1 | P2 | P3 | P4 | P5 | Avg | Errors | Notes |
|----------|----|----|----|----|----|----|--------|-------|
| 5.2 Range Close-Out | | | | | | | | |
| 5.3 Selective Reopen | | | | | | | | |
| 8.2 Segment Close-Out Decision | | | | | | | | |
| 9.2 Batch Close-Out from Weekly | | | | | | | | |
| 10.1 Room Type Close-Out Eval | | | | | | | | |
| 13.1 End-to-End Competitive Close-Out | | | | | | | | |

### Close-Out Decision Pattern Tracker

| Participant | First Data Checked | Close-Out Trigger | Dimensions Used (day/room/operator/board) | Preferred View | Checked Before Executing |
|-------------|-------------------|-------------------|-------------------------------------------|----------------|--------------------------|
| P1 | | | | | |
| P2 | | | | | |
| P3 | | | | | |
| P4 | | | | | |
| P5 | | | | | |
