# AI-Powered PPTX Generation - From Personal Hack to Product Vision

Presenter: Michael Ratke
Context: IPAI AI-Circle

## Slide 1 - Introduction

Before I start - this project is not related to any internal Exyte project. It's a side project that grew out of a real daily need.

Since I joined Exyte, I've had to create a lot of PowerPoint presentations - for architecture reviews, technical reviews, and various other meetings. Having a well-structured PPTX helps me stay on track and not forget anything during the meeting.

## Slide 2 - The Problem: I've Never Used PowerPoint

Here's the thing: I was born after PowerPoint had its peak popularity. I've actually never used it in my life. That doesn't mean I haven't done presentations - I've done plenty - but always in Canva or Google Slides. I never touched PPTX. And now I have to.

## Slide 3 - Research: Can Code Build PPTX?

Then I remembered - I'm an AI engineer. So I started researching. The first question in this kind of automation is: can we reach PPTX programmatically?

The answer came quickly. I found two well-maintained open-source frameworks:

1. python-pptx by Steve Canny (MIT license)
   - Provides deep control over the underlying XML structure
   - Lets you define slide design entirely in Python
   - Can both read and write PPTX files

2. pptxgenjs by Brent Ely (1.7M npm downloads)
   - Strong capabilities for building tables and charts
   - Popular in the JavaScript ecosystem
   - But it can only create PPTX - it cannot read or edit existing files
   - Keep this limitation in mind - it becomes relevant later
