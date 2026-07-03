import * as React from "react";
import { useState } from "react";
import styles from "./BpclMrDepartment.module.scss";
import type { IBpclMrDepartmentProps } from "./IBpclMrDepartmentProps";

//import ViewAllFavouriteLinks from "./ViewAllFavouriteLinks";
//import ViewAllDocuments from "./ViewAllDocuments";

import ViewAllEvents from "./ViewAllEvents";
import ViewAllAnnouncements from "./ViewAllAnnouncements";

import {
  Navbar,
  Nav,
  NavDropdown,
  Container,
  Button,
  Row,
  Col,
  Card,
  Modal,
} from "react-bootstrap";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap-icons/font/bootstrap-icons.css";
import "@fontsource/inter/400.css";
import "@fontsource/inter/500.css";
import "@fontsource/inter/600.css";
import "@fontsource/inter/700.css";
//import banner1 from "../assets/banner1.png";
import CombinedCalendar from "./CombinedCalendar";
import "react-calendar/dist/Calendar.css";
// import "slick-carousel/slick/slick.css";
// import "slick-carousel/slick/slick-theme.css";
import { Swiper, SwiperSlide } from "swiper/react";
import { Navigation } from "swiper/modules";

import "swiper/css";
import "swiper/css/navigation";
import news1 from "../assets/news1.png";
//import mission from "../assets/mission.png";
// import Slider from "react-slick";
// import { useState } from "react";

import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import BpclDepartmentService, {
  INavigationMenu,
  IWelcomeBanner,
  IVissionMission,
  IFavouriteLink,
  IEvent,
  ISharedDocument,
  IAnnouncement,
  IDiscussionBoard,
} from "../services/BpclDepartmentService";

const BpclMrDepartmentProps: React.FC<IBpclMrDepartmentProps> = (props) => {
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

  const [navigationMenu, setNavigationMenu] = React.useState<INavigationMenu[]>(
    [],
  );
  const [welcomeBanners, setWelcomeBanners] = React.useState<IWelcomeBanner[]>(
    [],
  );
  const [visionMission, setVisionMission] = React.useState<IVissionMission[]>(
    [],
  );
  const [favouriteLinks, setFavouriteLinks] = React.useState<IFavouriteLink[]>(
    [],
  );
  const [events, setEvents] = React.useState<IEvent[]>([]);
  const [documents, setDocuments] = React.useState<ISharedDocument[]>([]);
  const [announcements, setAnnouncements] = React.useState<IAnnouncement[]>([]);
  const [discussions, setDiscussions] = React.useState<IDiscussionBoard[]>([]);

  const [currentPage, setCurrentPage] = React.useState<
    "Home" | "FavouriteLinks" | "Events" | "Documents" | "Announcements"
  >("Home");

const siteTitle = props.context.pageContext.web.title;

  React.useEffect(() => {

  const siteHeader = document.getElementById("spSiteHeader");

  if (siteHeader) {
    siteHeader.style.display = "none";
  }

}, []);

  React.useEffect(() => {
    loadNavigationMenu();
    loadWelcomeBanners();
    loadVisionMission();
    loadFavouriteLinks();
    loadEvents();
    loadDocuments();
    loadAnnouncements();
    loadDiscussionBoard();
  }, []);

  const loadNavigationMenu = async (): Promise<void> => {
    try {
      const data = await service.getNavigationMenu();

//       const data = await service.getTermNavigation(
//     "BPCL-CPH",
//     "Award"
// ); 

      setNavigationMenu(data);

      console.log("Navigation Menu", data);
    } catch (error) {
      console.log(error);
    }
  };

  const loadWelcomeBanners = async (): Promise<void> => {
    try {
      const data = await service.getWelcomeBanners();

      setWelcomeBanners(data);

      console.log("Welcome Banner", data);
    } catch (error) {
      console.log(error);
    }
  };

  const slides = welcomeBanners;

  const [activeIndex, setActiveIndex] = React.useState(0);
  const [show, setShow] = useState(false);

  const handleClose = () => setShow(false);
  const handleShow = () => setShow(true);

  // AUTO SLIDER
  React.useEffect(() => {
    if (slides.length === 0) return;

    const interval = setInterval(() => {
      setActiveIndex((prev) => (prev === slides.length - 1 ? 0 : prev + 1));
    }, 4000);

    return () => clearInterval(interval);
  }, [slides]);

  // NEXT
  const nextSlide = () => {
    setActiveIndex((prev) => (prev === slides.length - 1 ? 0 : prev + 1));
  };
  // PREV
  const prevSlide = () => {
    setActiveIndex((prev) => (prev === 0 ? slides.length - 1 : prev - 1));
  };

  const activeSlide = slides.length > 0 ? slides[activeIndex] : undefined;

  React.useEffect(() => {
    if (activeIndex >= slides.length && slides.length > 0) {
      setActiveIndex(0);
    }
  }, [slides]);

  const loadVisionMission = async (): Promise<void> => {
    try {
      const data = await service.getVisionMission();

      setVisionMission(data);

      console.log("Vision Mission", data);
    } catch (error) {
      console.log(error);
    }
  };

  const mission = visionMission.find(
    (item) => item.Title.toLowerCase() === "mission",
  );

  const vision = visionMission.find(
    (item) => item.Title.toLowerCase() === "vision",
  );

  const loadFavouriteLinks = async (): Promise<void> => {
    try {
      const data = await service.getFavouriteLinks();

      setFavouriteLinks(data);

      console.log("Favourite Links", data);
    } catch (error) {
      console.log(error);
    }
  };

  const loadEvents = async (): Promise<void> => {
    try {
      const eventData = await service.getUpcomingEvents();

      setEvents(eventData);

      console.log("Events", eventData);
    } catch (error) {
      console.log(error);
    }
  };

  const loadDocuments = async () => {
    const data = await service.getRecentDocuments();

    setDocuments(data);
  };

  const getFileIcon = (fileType: string): string => {
    switch (fileType.toLowerCase()) {
      case "pdf":
        return "bi bi-file-earmark-pdf-fill text-danger";

      case "doc":
      case "docx":
        return "bi bi-file-earmark-word-fill text-primary";

      case "xls":
      case "xlsx":
        return "bi bi-file-earmark-excel-fill text-success";

      case "ppt":
      case "pptx":
        return "bi bi-file-earmark-ppt-fill text-warning";

      case "png":
      case "jpg":
      case "jpeg":
        return "bi bi-file-earmark-image-fill text-info";

      case "zip":
        return "bi bi-file-earmark-zip-fill text-secondary";

      default:
        return "bi bi-file-earmark-fill";
    }
  };

  const loadAnnouncements = async () => {
    const data = await service.getAnnouncements();

    setAnnouncements(data);
  };

  const loadDiscussionBoard = async () => {
    const data = await service.getDiscussionBoard();

    setDiscussions(data);
  };

  //   if (currentPage === "FavouriteLinks") {

  //     return (
  //         <ViewAllFavouriteLinks
  //             context={props.context}
  //             onBack={() => setCurrentPage("Home")}
  //         />
  //     );

  // }

  // if (currentPage === "Documents") {

  //     return (
  //         <ViewAllDocuments
  //             context={props.context}
  //             onBack={() => setCurrentPage("Home")}
  //         />
  //     );

  // }

  if (currentPage === "Events") {
    return (
      <ViewAllEvents
        context={props.context}
        onBack={() => setCurrentPage("Home")}
      />
    );
  }

  if (currentPage === "Announcements") {
    return (
      <ViewAllAnnouncements
        context={props.context}
        onBack={() => setCurrentPage("Home")}
      />
    );
  }

  return (
    <section className={styles.commonSectionstyle}>
      {/* NAVBAR */}
      <Navbar expand="lg" className={styles.navbarCustom}>
        <Container fluid>
          <Navbar.Brand className={styles.brandText}>
            {siteTitle}
          </Navbar.Brand>

          <Navbar.Toggle aria-controls="basic-navbar-nav" />

          <Navbar.Collapse id="basic-navbar-nav">
            <Nav className={`ms-auto ${styles.navWrapper}`}>
              {navigationMenu
                .filter((menu) => !menu.ParentIDId)
                .map((parent) => {
                  const children = navigationMenu.filter(
                    (child) => child.ParentIDId === parent.Id,
                  );

                  if (children.length > 0) {
                    return (
                      <NavDropdown
                        key={parent.Id}
                        title={parent.Title}
                        id={`menu-${parent.Id}`}
                      >
                        {children.map((child) => (
                          <NavDropdown.Item
                            key={child.Id}
                            href={child.RedirectURL?.Url}
                          >
                            {child.Title}
                          </NavDropdown.Item>
                        ))}
                      </NavDropdown>
                    );
                  }

                  return (
                    <Nav.Link key={parent.Id} href={parent.RedirectURL?.Url}>
                      {parent.Title}
                    </Nav.Link>
                  );
                })}
            </Nav>
          </Navbar.Collapse>
        </Container>
      </Navbar>

      {/* HERO SLIDER */}
      <div
        className={styles.heroSlide}
        style={{
          backgroundImage: activeSlide ? `url('${activeSlide.ImageUrl}')` : "",
        }}
      >
        <div className={styles.overlay}></div>

        {/* LEFT BUTTON */}
        <button className={styles.sliderBtnLeft} onClick={prevSlide}>
          <i className="bi bi-chevron-left"></i>
        </button>

        {/* RIGHT BUTTON */}
        <button className={styles.sliderBtnRight} onClick={nextSlide}>
          <i className="bi bi-chevron-right"></i>
        </button>

        {/* CONTENT */}
        <div className={styles.heroContent}>
          <h5 className={styles.welcomeText}>Welcome to</h5>

          <h1 className={styles.heroTitle}>
            {activeSlide?.WelcomeBannerTitle}
            <span className={styles.portalText}>Portal</span>
          </h1>

          <h3 className={styles.heroSubtitle}>
            Your one-stop platform for updates, resources and collaboration
          </h3>

          <p className={styles.heroDescription}>
            {activeSlide?.WelcomeBannerDesc}
          </p>

          {/* <div className="d-flex flex-column flex-sm-row flex-wrap gap-3 mt-4">
            <Button className={`${styles.primaryBtn} w-100 w-sm-auto`}>
              Explore More <i className="bi bi-arrow-right ms-2"></i>
            </Button>

            <Button className={`${styles.whiteBtn} w-100 w-sm-auto`}>
              Organogram <i className="bi bi-diagram-3 ms-2"></i>
            </Button>
          </div> */}
          <div className="d-flex flex-column flex-sm-row gap-3 mt-4">
            <Button className={`${styles.primaryBtn} w-100`}>
              Explore More <i className="bi bi-arrow-right ms-2"></i>
            </Button>

            <Button className={`${styles.whiteBtn} w-100`} onClick={handleShow}>
              Organogram <i className="bi bi-diagram-3 ms-2"></i>
            </Button>
          </div>
        </div>
      </div>

      {/* Mission & Vision , useful links */}
      <Container fluid className="p-4">
        <Row className="g-4 align-items-stretch h-100">
          {/* Left: Mission & Vision */}
          <Col xs={12} lg={8} className="d-flex">
            <div className={`${styles.mvSection} w-100`}>
              <Row className="mb-3">
                <Col xs={12} className={styles.titleWrap}>
                  <h2 className={styles.sectionTitle}>Mission & Vision</h2>
                </Col>
              </Row>

              <Row className="g-4 flex-grow-1">
                {/* Mission Card */}
                <Col xs={12} lg={6} className="d-flex">
                  <div className={`${styles.missionCard} w-100`}>
                    <div className={styles.cardHeader}>
                      <div
                      // className={`${styles.iconWrap} ${styles.missionIcon}`}
                      >
                        <img
                          src={mission?.AttachmentUrl}
                          alt="Mission"
                          width={50}
                          height={50}
                        />
                      </div>
                      <h3 className={styles.cardTitle}>{mission?.Title}</h3>
                    </div>

                    <div
                      className={styles.cardText}
                      dangerouslySetInnerHTML={{
                        __html: mission?.Description || "",
                      }}
                    />
                  </div>
                </Col>

                {/* Vision Card */}
                <Col xs={12} lg={6} className="d-flex">
                  <div className={`${styles.visionCard} w-100`}>
                    <div className={styles.cardHeader}>
                      <div
                      // className={`${styles.iconWrap} ${styles.visionIcon}`}
                      >
                        <img
                          src={vision?.AttachmentUrl}
                          alt="Vision"
                          width={50}
                          height={50}
                        />
                      </div>
                      <h3 className={styles.cardTitle}>{vision?.Title} </h3>
                    </div>

                    <div
                      className={styles.cardText}
                      dangerouslySetInnerHTML={{
                        __html: vision?.Description || "",
                      }}
                    />
                  </div>
                </Col>
              </Row>
            </div>
          </Col>

        <Col xs={12} lg={4} className="d-flex">
  <div className={`${styles.linksSection} w-100`}>

    {/* Header */}
    <div className={styles.newsHeader}>
      <h2>Useful Links</h2>

      {/* Show View All only when links are available */}
      {/* {favouriteLinks.length > 0 && (
        <a
          href={`${props.context.pageContext.web.absoluteUrl}/SitePages/ViewAllFavouriteLinks.aspx`}
          target="_blank"
          rel="noopener noreferrer"
        >
          View All
        </a>
      )} */}
    </div>

    {/* Content */}
    <div className={styles.linksContainer}>

      {favouriteLinks.length === 0 ? (

        <div
          className="d-flex justify-content-center align-items-center h-100"
          style={{
            minHeight: "320px",
            color: "#6c757d",
            fontSize: "16px",
            fontWeight: 500,
            textAlign: "center"
          }}
        >
          No Quick Links Available
        </div>

      ) : (

        favouriteLinks.map((link) => (

          <a
            key={link.Id}
            href={link.RedirectURL?.Url}
            target="_blank"
            rel="noopener noreferrer"
            className={`d-flex align-items-center justify-content-between mb-2 ${styles.linkCard}`}
            style={{ textDecoration: "none" }}
          >
            <div className="d-flex align-items-center gap-3">

              <div className={styles.iconWrapper}>
                <i className="bi bi-link"></i>
              </div>

              <div className={styles.textWrapper}>
                <span className={styles.linkDesc}>
                  {link.Title}
                </span>
              </div>

            </div>

            <div className={styles.arrowWrapper}>
              <i className="bi bi-chevron-right"></i>
            </div>

          </a>

        ))

      )}

    </div>

  </div>
</Col>
        </Row>
      </Container>
      {/* Calendar, events & Calendar */}
      <Container fluid className="p-4">
        <Row className="align-items-stretch h-100">
          <Col xs={12} lg={4} className="d-flex">
            <div className="w-100 d-flex flex-column h-100">
              <div className="flex-grow-1 h-100">
                <CombinedCalendar context={props.context} />
              </div>
            </div>
          </Col>
          {/* events */}
          <Col xs={12} lg={4}>
            <div className={`${styles.eventsSection} mt-4 mt-lg-0`}>
              {/* Header */}

              <div className={styles.newsHeader}>
                <h2>Events</h2>

                {events.length > 0 && (
                  <a
                    href="#"
                    className={styles.viewAllLink}
                    onClick={(e) => {
                      e.preventDefault();
                      setCurrentPage("Events");
                    }}
                  >
                    View All
                  </a>
                )}
              </div>
              {events.length === 0 ? (
                <div
                  className="d-flex justify-content-center align-items-center"
                  style={{
                    height: "320px",
                    color: "#6c757d",
                    fontSize: "16px",
                    fontWeight: 500, 
                  }}
                >
                  No Events Available
                </div>
              ) : (
                events.slice(0, 3).map((event, index) => {
                  const start = new Date(event.StartDate);
                  const end = event.EndDate ? new Date(event.EndDate) : start;

                  return (
                    <Card key={event.Id} className={styles.eventCard}>
                      <a
                        href={event.FileUrl || "#"}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{
                          textDecoration: "none",
                          color: "inherit",
                          display: "block",
                        }}
                        onClick={(e) => {
                          if (!event.FileUrl) {
                            e.preventDefault();
                          }
                        }}
                      >
                        <div
                          className={
                            index % 2 === 0
                              ? styles.leftBorder
                              : styles.leftBlueBorder
                          }
                        />

                        <Card.Body className={styles.eventCardBody}>
                          <Row className="align-items-center">
                            <Col xs="auto">
                              <div
                                className={
                                  index % 2 === 0
                                    ? styles.dateBox
                                    : styles.dateBlueBox
                                }
                              >
                                <div className={styles.day}>
                                  {start.getDate()}
                                </div>

                                <div className={styles.month}>
                                  {start.toLocaleString("default", {
                                    month: "short",
                                  })}
                                </div>
                              </div>
                            </Col>

                            <Col>
                              <h2 className={styles.eventTitle} title={event.Title}>
                                {event.Title}
                              </h2>

                              <div className={styles.eventTime}>
                                <i className="bi bi-clock" />

                                <span>
                                  {start.toLocaleTimeString([], {
                                    hour: "2-digit",
                                    minute: "2-digit",
                                  })}

                                  {" - "}

                                  {end.toLocaleTimeString([], {
                                    hour: "2-digit",
                                    minute: "2-digit",
                                  })}
                                </span>
                              </div>
                            </Col>
                          </Row>
                        </Card.Body>
                      </a>
                    </Card>
                  );
                })
              )} 
            </div>
          </Col>
          {/* shared documents */}
          <Col xs={12} lg={4} className="d-flex">
            <div
              className={`${styles.sharedDocumentsCard} w-100 d-flex flex-column h-100 mt-4 mt-lg-0`}
            >
              <div className={styles.newsHeader}>
                <h2>Shared Documents</h2>
                  {/* {documents.length > 0 && (
                <a
                  href="#"
                   className={styles.viewAllLink}
                  onClick={(e) => {
                    e.preventDefault();
                    setCurrentPage("Documents");
                  }}
                >
                  View All
                </a>
                  )} */}
              </div>

              <div className="flex-grow-1 ">
                {documents.length === 0 ? (
                  // <div className="text-center text-muted py-5">
                  <div
                  className="d-flex justify-content-center align-items-center"
                  style={{
                    height: "320px",
                    color: "#6c757d",
                    fontSize: "16px",
                    fontWeight: 500,
                  }}
                >
                    No Documents Available
                  </div>
                ) : (
                  documents.map((doc) => (
                    <Row
                      key={doc.Id}
                      className={styles.documentRow}
                      style={{ cursor: "pointer" }}
                      onClick={() => window.open(doc.FileUrl, "_blank")}
                    >
                      <Col xs="auto" className="d-flex align-items-center">
                        <div className={styles.fileIcon}>
                          <i className={`${getFileIcon(doc.FileType)}`} />
                        </div>
                      </Col>

                      <Col className="d-flex flex-column justify-content-center">
                        <div className={styles.documentTitle}>
                          {doc.FileName}
                        </div>

                        <div className={styles.documentDate}>
                          Uploaded on{" "}
                          {new Date(doc.Modified).toLocaleDateString("en-GB")}
                        </div>
                      </Col>

                      <Col xs="auto" className="d-flex align-items-center">
                        <a
                          href={doc.FileUrl}
                          target="_blank"
                          rel="noopener noreferrer"
                        >
                          <i className={`bi bi-eye-fill ${styles.eyeIcon}`} />
                        </a>
                      </Col>
                    </Row>
                  ))
                )}
              </div>
            </div>
          </Col>

          {/* RIGHT */}
        </Row>
      </Container>

      {/* <!-- ==========News slider  ========== --> */}
      <Container fluid className="p-4">
        <Row>
          {announcements.length > 0 && (
            <Col xs={12} lg={discussions.length > 0 ? 8 : 12}>
              <div className={styles.newsOuterSection}>
                {/* HEADER */}
                <div className={styles.newsHeader}>
                  <h2>News & Announcements</h2>
                  {/* <a
                    href={`${props.context.pageContext.web.absoluteUrl}/SitePages/ViewAllAnnouncements.aspx`}
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    View All
                  </a> */}
                  <a
                    href="#"
                    className={styles.viewAllLink}
                    onClick={(e) => {
                      e.preventDefault();
                      setCurrentPage("Announcements");
                    }}
                  >
                    View All
                  </a>
                </div>

                {/* SWIPER WRAPPER */}
                <div className={`${styles.newsSliderWrapper} mt-3`}>
                  {/* LEFT BUTTON */}
                  <button
                    type="button"
                    className={`news-prev ${styles.swiperNav} ${styles.swiperPrev}`}
                  >
                    <i className="bi bi-chevron-left" />
                  </button>

                  {/* SWIPER */}
                  <Swiper
                    modules={[Navigation]}
                    spaceBetween={20}
                    slidesPerView={3}
                    slidesPerGroup={1}
                    navigation={{
                      prevEl: ".news-prev",
                      nextEl: ".news-next",
                    }}
                    breakpoints={{
                      0: { slidesPerView: 1 },
                      768: { slidesPerView: 2 },
                      1200: { slidesPerView: 3 },
                    }}
                  >
                    {announcements.map((item) => (
                      <SwiperSlide key={item.Id}>
                        <a
                          href={item.FileUrl}
                          target="_blank"
                          rel="noopener noreferrer"
                          className={styles.newsCard}
                        >
                          <div className={styles.cardImageWrapper}>
                            <img
                              src={
                                item.ImageUrl || news1
                              }
                              alt={item.Title}
                              className={styles.cardImage}
                            />

                            <div className={styles.cardOverlay}>
                              <h6
                                className={styles.newscardTitle}
                                title={item.Title}
                              >
                                {item.Title}{" "}
                              </h6>

                              <div className={styles.cardDate}>
                                <i className="bi bi-calendar3" />
                                <span>
                                  {item.PublishedDate
                                    ? new Date(
                                        item.PublishedDate,
                                      ).toLocaleDateString("en-GB")
                                    : ""}
                                </span>
                              </div>
                            </div>
                          </div>
                        </a>
                      </SwiperSlide>
                    ))}
                  </Swiper>

                  {/* RIGHT BUTTON */}
                  <button
                    type="button"
                    className={`news-next ${styles.swiperNav} ${styles.swiperNext}`}
                  >
                    <i className="bi bi-chevron-right" />
                  </button>
                </div>
              </div>
            </Col>
          )}
          {discussions.length > 0 && (
            <Col xs={12} lg={announcements.length > 0 ? 4 : 12}>
              <div className={styles.discussionBoardSection}>
                {/* Header */}
                <div className={styles.newsHeader}>
                  <h2>General Discussion Board</h2>

                  <a
                    href={`${props.context.pageContext.web.absoluteUrl}/Lists/MR_SL_DiscussionBoard/AllItems.aspx`}
                    className={styles.viewAllLink} 
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    View All
                  </a>
                </div>

                {discussions.slice(0, 3).map((item) => (
                  <div
                    key={item.Id}
                    className={styles.discussionCard}
                    style={{ cursor: "pointer" }}
                    onClick={() => {
                      if (item.RedirectURL?.Url) {
                        window.open(item.RedirectURL.Url, "_blank");
                      }
                    }}
                  >
                    <div className={styles.statusBar} />

                    <div className={styles.content}>
                      <div className="row align-items-center">
                        <div className="col-12">
                          <h3 className={styles.title} title={item.Title}>
                            {item.Title}
                          </h3>

                          <div className={styles.metadata}>
                            Latest Reply :
                            <span className={styles.author}>
                              {item.LatestReplyBy?.Title}
                            </span>
                            <span className={styles.separator}>•</span>
                            <span className={styles.date}>
                              {new Date(item.ReplyDate).toLocaleDateString(
                                "en-GB",
                              )}
                            </span>
                            <span className={styles.time}>
                              {new Date(item.ReplyDate).toLocaleTimeString([], {
                                hour: "2-digit",
                                minute: "2-digit",
                              })}
                            </span>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </Col>
          )}{" "}
        </Row>
      </Container>

      {/* Organogram Modal  */}
      <Modal show={show} onHide={handleClose} centered size="xl">
        <Modal.Header closeButton>
          <Modal.Title className={styles.modalTitle}>Organogram</Modal.Title>
        </Modal.Header>

        <Modal.Body className={`p-3 ${styles.modalBody}`}>
          {/* <iframe
            src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d248755.79475951716!2d80.04385848151072!3d13.047807814214304!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x3a5265ea4f7d3361%3A0x6e61a70b6863d433!2sChennai%2C%20Tamil%20Nadu!5e0!3m2!1sen!2sin!4v1782449802845!5m2!1sen!2sin"
            title=""
            width="100%"
            height="600"
            frameBorder="0"
            style={{ border: 0 }}
          /> */}
          <iframe
            title="Organogram"
            src="https://bharatpetroleum.sharepoint.com/sites/dev-mumbai-refinery-advisory-services/SitePages/Organogram.aspx?env=Embedded"
            width="100%"
            height="900"
            style={{
              border: "none",
              borderRadius: "12px",
            }}
          />
        </Modal.Body>
      </Modal>
    </section>
  );
};

export default BpclMrDepartmentProps;
