import * as React from "react";
import Calendar from "react-calendar";
import moment from "moment";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Modal } from "react-bootstrap";
interface ICombinedCalendarProps {
  context: WebPartContext;
}

interface IEventItem {
  id?: number;
  title: string;
  date: Date;
  type: "sp" | "teams";
  color: string;
  start?: Date;
  end?: Date;
  joinUrl?: string;
  description?: string;
}

interface ICombinedCalendarState {
  events: IEventItem[];
  selectedDate: Date;
  hoverEvents: IEventItem[];

  tooltipTop: number;
  tooltipLeft: number;

  isMobile: boolean;
  showModal: boolean;
  selectedEvents: IEventItem[];
}

export default class CombinedCalendar extends React.Component<
  ICombinedCalendarProps,
  ICombinedCalendarState
> {
  constructor(props: ICombinedCalendarProps) {
    super(props);

    this.state = {
      events: [],
      selectedDate: new Date(),
      hoverEvents: [],
      tooltipTop: 0,
      tooltipLeft: 0,
      isMobile: false,
      showModal: false,
      selectedEvents: [],
    };
  }

  // ---------------- MOBILE ----------------
  componentDidMount(): void {
    this.setState({ isMobile: window.innerWidth <= 768 });
    window.addEventListener("resize", this.handleResize);

    this.loadSharePointEvents();
    this.loadTeamsMeetings();
  }

  componentWillUnmount(): void {
    window.removeEventListener("resize", this.handleResize);
  }

  private handleResize = (): void => {
    this.setState({ isMobile: window.innerWidth <= 768 });
  };
 
 private async loadSharePointEvents(): Promise<void> {

  const siteUrl = this.props.context.pageContext.web.absoluteUrl;

  const today = new Date().toISOString();

  const response = await fetch(
    `${siteUrl}/_api/web/lists/getbytitle('MR_SL_Events')/items?$select=Id,Title,Description,StartDate,EndDate,IsActive&$filter=IsActive eq 1 and EndDate ge datetime'${today}'&$orderby=StartDate asc`,
    {
      headers: {
        Accept: "application/json;odata=nometadata",
      },
    }
  );

  const data = await response.json();

  const spEvents: IEventItem[] = data.value.map((item: any) => ({

    id: item.Id,

    title: item.Title,

    description: item.Description,

    date: new Date(item.StartDate),

    start: new Date(item.StartDate),

    end: new Date(item.EndDate),

    type: "sp",

    color: "#FFA300",

  }));

  this.setState((p) => ({
    events: [...p.events, ...spEvents],
  }));
}

  
  private async loadTeamsMeetings(): Promise<void> {
    const client = await this.props.context.msGraphClientFactory.getClient("3");

    const startDate = new Date();
    startDate.setMonth(startDate.getMonth() - 1);

    const endDate = new Date();
    endDate.setMonth(endDate.getMonth() + 2);

    const response = await client
      .api("/me/calendarView")
      .query({
        startDateTime: startDate.toISOString(),
        endDateTime: endDate.toISOString(),
      })
      .top(100)
      .get();

    

    const teamEvents: IEventItem[] = response.value

      // REMOVE CANCELLED EVENTS

      .filter((item: any) => !item.isCancelled)

      .map((item: any) => ({
        title: item.subject,

        date: moment

          .utc(item.start.dateTime)

          .local()

          .toDate(),

        start: moment

          .utc(item.start.dateTime)

          .local()

          .toDate(),

        end: moment

          .utc(item.end.dateTime)

          .local()

          .toDate(),

        type: "teams",

        color: "#0ea5e9",

        joinUrl: item.onlineMeeting?.joinUrl,
      }));

    this.setState((p) => ({ events: [...p.events, ...teamEvents] }));
  }

  // ---------------- EVENTS ----------------
  // private getEvents = (date: Date): IEventItem[] => {
  //   return this.state.events.filter(
  //     (e) =>
  //       moment(e.date).format("YYYY-MM-DD") ===
  //       moment(date).format("YYYY-MM-DD"),
  //   );
  // };

  private getEvents = (date: Date): IEventItem[] => {

  return this.state.events.filter((e) => {

    if (!e.start) return false;

    const start = moment(e.start).startOf("day");

    const end = e.end
      ? moment(e.end).startOf("day")
      : start;

    const current = moment(date).startOf("day");
    
    return current.isBetween(start, end, undefined, "[]");

  });

};

  // ---------------- TILE ----------------
  private renderTileContent = ({ date, view }: any): React.ReactNode => {
    if (view !== "month") return null;

    const events = this.getEvents(date);

    const hasSP = events.some((e) => e.type === "sp");
    const hasTeams = events.some((e) => e.type === "teams");

    // const isToday =
    //   moment(date).format("YYYY-MM-DD") === moment().format("YYYY-MM-DD");

    return (
      <div
        onMouseEnter={(e) => this.onHover(date, e.currentTarget)}
        onMouseLeave={this.onLeave}
        onClick={() => this.onDayClick(date)}
        style={{
          display: "flex",
          justifyContent: "center",
          gap: "4px",
          marginTop: "2px",
          cursor: "pointer",
          padding: "2px",
          borderRadius: "6px",
          background: "transparent",
          border: "1px solid transparent",
        }}
      >
        {hasSP && (
          <div
            style={{
              width: "7px",
              height: "7px",
              borderRadius: "50%",
              background: "#FFA300",
            }}
          />
        )}

        {hasTeams && (
          <div
            style={{
              width: "7px",
              height: "7px",
              borderRadius: "50%",
              background: "#009EE9",
            }}
          />
        )}
      </div>
    );
  };

  // ---------------- HOVER ----------------
  private onHover = (date: Date, element: HTMLElement): void => {
    if (this.state.isMobile) return;

    const events = this.getEvents(date);
    if (events.length === 0) return;

    const rect = element.getBoundingClientRect();

    this.setState({
      hoverEvents: events,
      tooltipTop: rect.top - 10,
      tooltipLeft: rect.left + rect.width / 2,
    });
  };

  private onLeave = (): void => {
    this.setState({
      hoverEvents: [],
    });
  };

  // ---------------- CLICK ----------------
  // private onDayClick = (date: Date): void => {
  //   this.setState({
  //     selectedDate: date,
  //   });
  // };

  private onDayClick = (date: Date): void => {
    const events = this.getEvents(date);

    this.setState({
      selectedDate: date,

      selectedEvents: events,

      showModal: events.length > 0,
    });
  };

 

  // ---------------- RENDER ----------------
  public render(): React.ReactElement<ICombinedCalendarProps> {

    return (
      <div style={{ display: "flex", gap: "16px", position: "relative" }}>
        {/* CALENDAR */}
        <div style={{ flex: 2 }}>
          <div
            style={{
              position: "relative",
              border: "1px solid #f1f3f6",
              borderRadius: "10px",
              padding: "30px",
              background: "#fff",
            }}
          >
            {/* LEGEND */}
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                gap: "16px",
                alignItems: "center",
                padding: "6px 0",
                fontSize: "12px",
                marginBottom: "8px",
              }}
            >
              <div
                style={{ display: "flex", alignItems: "center", gap: "5px" }}
              >
                <div
                  style={{
                    width: "8px",
                    height: "8px",
                    borderRadius: "50%",
                    background: "#FFA300",
                  }}
                />
                Events
              </div>

              <div
                style={{ display: "flex", alignItems: "center", gap: "5px" }}
              >
                <div
                  style={{
                    width: "8px",
                    height: "8px",
                    borderRadius: "50%",
                    background: "#009EE9",
                  }}
                />
                Teams Meeting
              </div>
            </div>

            {/* CALENDAR */}
            <Calendar
              value={this.state.selectedDate}
              onClickDay={this.onDayClick}
              tileContent={this.renderTileContent}
            />
          </div>
        </div>

        <Modal
          show={this.state.showModal}
          onHide={() =>
            this.setState({
              showModal: false,
            })
          }
          centered
          size="lg"
        >
          <Modal.Header closeButton>
            <Modal.Title>
              {moment(this.state.selectedDate).format("DD MMMM YYYY")}
            </Modal.Title>
          </Modal.Header>

          <Modal.Body>
            {this.state.selectedEvents.length === 0 && (
              <div className="text-center text-muted py-5">
                No events available.
              </div>
            )}
            {this.state.selectedEvents.map((e, index) => {

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const meetingDate = new Date(e.start!);
  meetingDate.setHours(0, 0, 0, 0);

  const showJoinButton =
    e.type === "teams" &&
    !!e.joinUrl &&
    meetingDate >= today;

  return (
              <div
                key={index}
                style={{
                  borderBottom: "1px solid #eee",
                  paddingBottom: 15,
                  marginBottom: 15,
                }}
              >
                {/* <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 10,
                  }}
                >
                   <div
                    style={{
                      width: 10,
                      height: 10,
                      borderRadius: "50%",
                      background: e.type === "teams" ? "#009EE9" : "#FFA300",
                    }}
                  />

                  <strong>{e.title}</strong>
                </div>  */}

                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    marginBottom: 8,
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 8,
                    }}
                  >
                    <div
                      style={{
                        width: 10,
                        height: 10,
                        borderRadius: "50%",
                        background: e.type === "teams" ? "#009EE9" : "#FFA300",
                      }}
                    />

                    <strong>{e.title}</strong>
                  </div>
                </div>
                <div
                  style={{
                    marginTop: 5,
                    color: "#666",
                  }}
                >
                  {moment(e.start).format("hh:mm A")}

                  {" - "}

                  {moment(e.end).format("hh:mm A")}
                </div>

                {e.description && (
                  <div
                    style={{
                      marginTop: 8,
                      color: "#666",
                      fontSize: "13px",
                    }}
                  >
                    {e.description}
                  </div>
                )}

                {/* {e.type === "teams" && e.joinUrl && (
                  <a
                    href={e.joinUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="btn btn-primary btn-sm mt-2"
                  >
                    Join Meeting
                  </a>
                )} */}

                {showJoinButton && (
  <a
    href={e.joinUrl}
    target="_blank"
    rel="noopener noreferrer"
    className="btn btn-primary btn-sm mt-2"
  >
    Join Meeting
  </a>
)}

</div>

);

})}
          </Modal.Body>
        </Modal>

        {/* TOOLTIP */}
         {this.state.hoverEvents.length > 0 && (
          <div
            style={{
              position: "fixed",
              top: this.state.tooltipTop,
              left: this.state.tooltipLeft,
              transform: "translate(-50%, -100%)",
              background: "#fff",
              borderRadius: "12px",
              boxShadow: "0 12px 30px rgba(0,0,0,0.15)",
              border: "1px solid #e5e7eb",
              padding: "12px",
              width: "260px",
              zIndex: 99999,
            }}
          >
            <div style={{ fontWeight: 600, marginBottom: "8px" }}>Events</div>
 
            {this.state.hoverEvents.map((e, i) => (
              <div key={i} style={{ padding: "4px 0" }}>
                {e.title}
              </div>
            ))}
          </div>
        )} 

         {this.state.hoverEvents.length > 0 && (
          <div
            style={{
              position: "fixed",
              top: this.state.tooltipTop,
              left: this.state.tooltipLeft,
              transform: "translate(-50%, -100%)",
              background: "rgba(255,255,255,0.95)",
              backdropFilter: "blur(10px)",
              WebkitBackdropFilter: "blur(10px)",
              borderRadius: "14px",
              boxShadow: "0 18px 40px rgba(0,0,0,0.18)",
              border: "1px solid rgba(0,0,0,0.08)",
              padding: "12px 14px",
              width: "280px",
              zIndex: 99999,
              animation: "fadeIn 0.15s ease-in-out",
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                marginBottom: "10px",
              }}
            >
              <div
                style={{
                  fontWeight: 600,
                  fontSize: "13px",
                  color: "#111827",
                  display: "flex",
                  alignItems: "center",
                  gap: "6px",
                }}
              >
                 Events & Teams
              </div>

              <div
                style={{
                  fontSize: "11px",
                  color: "#6b7280",
                }}
              >
                {this.state.hoverEvents.length}
              </div>
            </div>

            <div
              style={{ display: "flex", flexDirection: "column", gap: "8px" }}
            >
              {this.state.hoverEvents.map((e, i) => (
                <div
                  key={i}
                  style={{
                    display: "flex",
                    alignItems: "flex-start",
                    gap: "10px",
                    padding: "8px 10px",
                    borderRadius: "10px",
                    background: "#f9fafb",
                    border: "1px solid #f1f5f9",
                    cursor: "default",
                  }}
                >
                  <div
                    style={{
                      width: "8px",
                      height: "8px",
                      borderRadius: "50%",
                      marginTop: "5px",
                      background: e.type === "sp" ? "#FFA300" : "#009EE9",
                      flexShrink: 0,
                    }}
                  />

                  <div style={{ flex: 1 }}>
                    <div
                      style={{
                        fontSize: "12.5px",
                        fontWeight: 500,
                        color: "#111827",
                        lineHeight: "1.3",
                      }}
                    >
                      {e.title}
                    </div>

                    <div
                      style={{
                        fontSize: "11px",
                        color: "#6b7280",
                        marginTop: "2px",
                      }}
                    >
                      {e.start ? moment(e.start).format("hh:mm A") : ""}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}  
      </div>
    );
  }
}
