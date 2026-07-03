import * as React from 'react';
import styles from './ViewAllPagesCommon.module.scss';
//import '@fontsource/inter';
import '@fontsource/inter/500.css';
import '@fontsource/inter/600.css';
import '@fontsource/inter/700.css';
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap-icons/font/bootstrap-icons.css';

 
import {
  Row,
  Col,
  Form, 
  Button,
  Pagination
} from 'react-bootstrap';

import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp";

import BpclDepartmentService, {
  IEvent,
} from "../services/BpclDepartmentService";



import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IViewAllEventsProps {
  context: WebPartContext;
  onBack: () => void;
}



const ViewAllEvents: React.FC<IViewAllEventsProps> = (props) => {
    
    const sp: SPFI = React.useMemo(() => {
  return spfi().using(SPFx(props.context));
}, [props.context]);

// const service = React.useMemo(() => {
//   return new BpclDepartmentService(sp);
// }, [sp]);

 const service = React.useMemo(() => {
    return new BpclDepartmentService(
        sp,
        props.context
    );
}, [sp, props.context]);

const [events, setEvents] = React.useState<IEvent[]>([]);

const [filteredEvents, setFilteredEvents] =
React.useState<IEvent[]>([]);

const [fromDate, setFromDate] =
React.useState("");

const [toDate, setToDate] =
React.useState("");

const [currentPage, setCurrentPage] =
React.useState(1);

const [pageSize, setPageSize] =
React.useState(10);



const loadEvents = async () => {

const data = await service.getAllEvents();

    setEvents(data);

    setFilteredEvents(data);

};

React.useEffect(() => {
    loadEvents();
}, [service]);

const handleSearch = () => {

    let data = [...events];

    if (fromDate) {

        const from = new Date(fromDate);
        from.setHours(0, 0, 0, 0);

        data = data.filter(item => {

            const start = new Date(item.StartDate);
            start.setHours(0, 0, 0, 0);

            return start >= from;

        });

    }

    if (toDate) {

        const to = new Date(toDate);
        to.setHours(23, 59, 59, 999);

        data = data.filter(item => {

            const start = new Date(item.StartDate);

            return start <= to;

        });

    }

    setFilteredEvents(data);

    setCurrentPage(1);

}; 

const totalPages =
    Math.ceil(filteredEvents.length / pageSize);

const pagedEvents = filteredEvents.slice(

    (currentPage - 1) * pageSize,

    currentPage * pageSize

);

  return (
    <div className={styles.pageContainer}>

      {/* Banner */}

  <div className={styles.banner}>
  <button
  className={styles.backIcon}
//   onClick={() => window.history.back()}
onClick={props.onBack}
>
  <i className="bi bi-chevron-left"></i>
</button>
  <div className={styles.bannerIcon}>
    <i className="bi bi-calendar-event-fill"></i>
  </div>
 
  <div>
    <h3>Events</h3>
    <p>
      Discover upcoming events, workshops, meetings, celebrations, and activities happening across the organization.
    </p>
  </div>
 
</div>
      
<div className={styles.cardContainer}>

          <Row className="gx-3 mb-5 align-items-end">

            <Col xl={4} lg={4} md={6} sm={12}>
              <Form.Label className={styles.formLabel}>From</Form.Label>
              <div className={styles.dateBox}>
<Form.Control
type="date"
value={fromDate}
onChange={(e)=>setFromDate(e.target.value)}
/>              </div>
            </Col>

            <Col xl={4} lg={4} md={6} sm={12}>
              <Form.Label className={styles.formLabel}>To</Form.Label>
              <div className={styles.dateBox}>
<Form.Control
type="date"
value={toDate}
onChange={(e)=>setToDate(e.target.value)}
/>              </div>
            </Col>

            <Col xl={2} lg={4} md={12} sm={12}>
<Button
className={`w-100 mt-sm-3 ${styles.searchBtn}`}
onClick={handleSearch}
>                <i className="bi bi-search me-2"></i>
                Search
              </Button>
            </Col>

          </Row>
{/* </Card.Body></Card> */}

   <Row className="gx-3">

    {pagedEvents.length === 0 ? (

<div className="text-center py-5">

No Events Found

</div>

) : (



pagedEvents.map((card) => {

const start = new Date(card.StartDate);

const end = card.EndDate
  ? new Date(card.EndDate)
  : start;

return (

<Col
  xl={3}
  lg={3}
  md={6}
  sm={12}
  className="mb-4 d-flex"
  key={card.Id}
>

<a
  href={card.FileUrl || "#"}
  target="_blank"
  rel="noopener noreferrer"
  className={`card h-100 text-decoration-none ${styles.eventCard}`}
  style={{ color: "inherit" }}
  onClick={(e) => {
    if (!card.FileUrl) {
      e.preventDefault();
    }
  }}
>

<div className={styles.eventIcon}>
  <i className="bi bi-calendar-event-fill"></i>
</div>

<h5 className={styles.eventTitle}>
  {card.Title}
</h5>

<div className={styles.eventDateRow}>

  <i className="bi bi-calendar-event"></i>

  <span className={styles.dateText}>

    <strong>
      {start.toLocaleDateString("en-GB")}
    </strong>

    <span className={styles.dateSeparator}>
      &nbsp;-&nbsp;
    </span>

    <strong>
      {end.toLocaleDateString("en-GB")}
    </strong>

  </span>

</div>

</a>

</Col>

);

})

)}


</Row>

          {/* Footer */}

         <div className={styles.paginationSection}>

  <div className={styles.leftPagination}>

    <span className={styles.pageInfo}>
Page {currentPage} of {totalPages}
    </span>

<Form.Select
value={pageSize}
onChange={(e) => {
    setPageSize(Number(e.target.value));
    setCurrentPage(1);
}}
>

<option value={10}>10 per page</option>

<option value={20}>20 per page</option>

<option value={50}>50 per page</option>

</Form.Select>

  </div>

 <Pagination className="mb-0 justify-content-end flex-wrap">

<Pagination.First
disabled={currentPage===1}
onClick={()=>setCurrentPage(1)}
/>

<Pagination.Prev
disabled={currentPage===1}
onClick={()=>setCurrentPage(currentPage-1)}
/>

{[...Array(totalPages)].map((_,index)=>(

<Pagination.Item

key={index}

active={currentPage===index+1}

onClick={()=>setCurrentPage(index+1)}

>

{index+1}

</Pagination.Item>

))}

<Pagination.Next
disabled={currentPage===totalPages}
onClick={()=>setCurrentPage(currentPage+1)}
/>

<Pagination.Last
disabled={currentPage===totalPages}
onClick={()=>setCurrentPage(totalPages)}
/>

</Pagination>

</div>

       
</div>


    </div>

   
  );
};

export default ViewAllEvents;