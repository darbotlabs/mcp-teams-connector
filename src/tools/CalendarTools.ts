import { Client } from '@microsoft/microsoft-graph-client';
import { Event, Attendee } from '@microsoft/microsoft-graph-types';

export class CalendarTools {
  constructor(private graphClient: Client) {}

  async createMeeting(params: {
    subject: string;
    startTime: string;
    endTime: string;
    attendees?: string[];
    body?: string;
    location?: string;
    isOnline?: boolean;
  }): Promise<any> {
    const event: any = {
      subject: params.subject,
      start: {
        dateTime: params.startTime,
        timeZone: 'UTC'
      },
      end: {
        dateTime: params.endTime,
        timeZone: 'UTC'
      },
      attendees: params.attendees?.map(email => ({
        emailAddress: { address: email },
        type: 'required'
      })) || [],
      body: {
        content: params.body || '',
        contentType: 'text'
      },
      isOnlineMeeting: params.isOnline !== false, // Default to true
      onlineMeetingProvider: params.isOnline !== false ? 'teamsForBusiness' : undefined
    };

    // Add location if provided
    if (params.location) {
      event.location = { displayName: params.location };
    }

    try {
      console.log(`Creating meeting: ${params.subject}`);
      const createdEvent = await this.graphClient
        .api('/me/events')
        .post(event);

      const joinUrl = createdEvent.onlineMeeting?.joinUrl;
      const meetingInfo = [
        `Meeting "${params.subject}" created successfully!`,
        `Meeting ID: ${createdEvent.id}`,
        `Start: ${new Date(params.startTime).toLocaleString()}`,
        `End: ${new Date(params.endTime).toLocaleString()}`
      ];

      if (joinUrl) {
        meetingInfo.push(`Join URL: ${joinUrl}`);
      }

      if (params.attendees && params.attendees.length > 0) {
        meetingInfo.push(`Attendees: ${params.attendees.join(', ')}`);
      }

      return {
        content: [{
          type: 'text',
          text: meetingInfo.join('\n')
        }]
      };
    } catch (error: any) {
      console.error('Failed to create meeting:', error);
      return {
        content: [{
          type: 'text',
          text: `Failed to create meeting: ${error.message || 'Unknown error'}`
        }],
        isError: true
      };
    }
  }

  async findAvailability(params: {
    attendees: string[];
    startDate: string;
    endDate: string;
    duration: number;
  }): Promise<any> {
    try {
      console.log(`Checking availability for ${params.attendees.length} attendees`);
      
      const scheduleInformation = {
        schedules: params.attendees,
        startTime: {
          dateTime: params.startDate,
          timeZone: 'UTC'
        },
        endTime: {
          dateTime: params.endDate,
          timeZone: 'UTC'
        },
        availabilityViewInterval: 30
      };

      const response = await this.graphClient
        .api('/me/calendar/getSchedule')
        .post(scheduleInformation);

      // Process availability data and find suitable slots
      const availableSlots = this.findMeetingSlots(
        response.value,
        params.duration,
        new Date(params.startDate),
        new Date(params.endDate),
        params.attendees
      );

      const resultText = availableSlots.length > 0
        ? `Found ${availableSlots.length} available time slots:\n${
            availableSlots.map(slot => 
              `• ${slot.start.toLocaleString()} - ${slot.end.toLocaleString()} (${slot.availableAttendees}/${params.attendees.length} available)`
            ).join('\n')
          }`
        : `No common availability found for all attendees in the specified time range. Consider:\n` +
          `- Expanding the date range\n` +
          `- Reducing the meeting duration (currently ${params.duration} minutes)\n` +
          `- Checking individual attendee schedules`;

      return {
        content: [{
          type: 'text',
          text: resultText
        }]
      };
    } catch (error: any) {
      console.error('Failed to check availability:', error);
      return {
        content: [{
          type: 'text',
          text: `Failed to check availability: ${error.message || 'Unknown error'}`
        }],
        isError: true
      };
    }
  }

  private findMeetingSlots(
    schedules: any[], 
    duration: number, 
    startDate: Date, 
    endDate: Date,
    attendees: string[]
  ): Array<{start: Date; end: Date; availableAttendees: number}> {
    const slots: Array<{start: Date; end: Date; availableAttendees: number}> = [];
    
    // Create 30-minute time slots
    const timeSlot = new Date(startDate);
    const durationMs = duration * 60 * 1000;
    
    while (timeSlot < endDate) {
      const slotEnd = new Date(timeSlot.getTime() + durationMs);
      
      // Don't create slots that extend beyond the end date
      if (slotEnd > endDate) {
        break;
      }
      
      // Skip slots outside business hours (9 AM - 5 PM)
      const hour = timeSlot.getHours();
      if (hour < 9 || hour >= 17) {
        timeSlot.setMinutes(timeSlot.getMinutes() + 30);
        continue;
      }
      
      // Skip weekends
      const dayOfWeek = timeSlot.getDay();
      if (dayOfWeek === 0 || dayOfWeek === 6) {
        timeSlot.setMinutes(timeSlot.getMinutes() + 30);
        continue;
      }
      
      let availableCount = 0;
      
      // Check availability for each attendee
      for (let i = 0; i < schedules.length; i++) {
        const schedule = schedules[i];
        if (this.isTimeFree(schedule, timeSlot, duration)) {
          availableCount++;
        }
      }
      
      // Include slots where at least half the attendees are available
      if (availableCount >= Math.ceil(attendees.length / 2)) {
        slots.push({
          start: new Date(timeSlot),
          end: new Date(slotEnd),
          availableAttendees: availableCount
        });
      }
      
      // Move to next 30-minute slot
      timeSlot.setMinutes(timeSlot.getMinutes() + 30);
    }
    
    // Sort by most available attendees first, then by time
    return slots
      .sort((a, b) => {
        if (b.availableAttendees !== a.availableAttendees) {
          return b.availableAttendees - a.availableAttendees;
        }
        return a.start.getTime() - b.start.getTime();
      })
      .slice(0, 10); // Return top 10 slots
  }

  private isTimeFree(schedule: any, startTime: Date, duration: number): boolean {
    if (!schedule.freeBusyViewInfo) {
      return true; // Assume free if no data
    }
    
    const endTime = new Date(startTime.getTime() + duration * 60 * 1000);
    
    // Check if any busy periods conflict with our time slot
    for (const busyTime of schedule.busyTimes || []) {
      const busyStart = new Date(busyTime.start.dateTime);
      const busyEnd = new Date(busyTime.end.dateTime);
      
      // Check for overlap
      if (startTime < busyEnd && endTime > busyStart) {
        return false;
      }
    }
    
    return true;
  }

  async listUpcomingMeetings(params: {
    days?: number;
    includeDetails?: boolean;
  } = {}): Promise<any> {
    const days = params.days || 7;
    const startDateTime = new Date().toISOString();
    const endDateTime = new Date();
    endDateTime.setDate(endDateTime.getDate() + days);

    try {
      console.log(`Retrieving meetings for the next ${days} days`);
      
      const events = await this.graphClient
        .api('/me/calendarView')
        .query({
          startDateTime,
          endDateTime: endDateTime.toISOString(),
          $orderby: 'start/dateTime',
          $select: params.includeDetails 
            ? 'subject,start,end,location,attendees,onlineMeeting,body,organizer,webLink' 
            : 'subject,start,end,location,onlineMeeting'
        })
        .get();

      if (!events.value || events.value.length === 0) {
        return {
          content: [{
            type: 'text',
            text: `No upcoming meetings found for the next ${days} days.`
          }]
        };
      }

      const meetingList = events.value.map((event: Event, index: number) => {
        const start = new Date(event.start!.dateTime!);
        const end = new Date(event.end!.dateTime!);
        
        let text = `${index + 1}. ${event.subject}`;
        text += `\n   📅 ${start.toLocaleDateString()} ${start.toLocaleTimeString()} - ${end.toLocaleTimeString()}`;
        
        if (event.location?.displayName) {
          text += `\n   📍 ${event.location.displayName}`;
        }
        
        if (event.onlineMeeting?.joinUrl) {
          text += `\n   💻 Teams Meeting`;
        }
        
        if (params.includeDetails) {
          if (event.attendees && event.attendees.length > 0) {
            const attendeeNames = event.attendees
              .map((a: any) => a.emailAddress?.name || a.emailAddress?.address)
              .filter(Boolean)
              .join(', ');
            text += `\n   👥 Attendees: ${attendeeNames}`;
          }
          
          if (event.organizer?.emailAddress?.name) {
            text += `\n   👤 Organizer: ${event.organizer.emailAddress.name}`;
          }
          
          if (event.body?.content && event.body.content.trim()) {
            const bodyText = event.body.content
              .replace(/<[^>]*>/g, '') // Remove HTML tags
              .substring(0, 100)
              .trim();
            if (bodyText) {
              text += `\n   📝 ${bodyText}${event.body.content.length > 100 ? '...' : ''}`;
            }
          }
        }
        
        return text;
      }).join('\n\n');

      return {
        content: [{
          type: 'text',
          text: `Upcoming meetings (next ${days} days):\n\n${meetingList}`
        }]
      };
    } catch (error: any) {
      console.error('Failed to retrieve meetings:', error);
      return {
        content: [{
          type: 'text',
          text: `Failed to retrieve meetings: ${error.message || 'Unknown error'}`
        }],
        isError: true
      };
    }
  }
}