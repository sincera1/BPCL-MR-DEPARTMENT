import * as React from "react";
import { useState } from "react";
import styles from "./BpclMrDepartment.module.scss";
import type { IBpclMrDepartmentProps } from "./IBpclMrDepartmentProps";

//import ViewAllFavouriteLinks from "./ViewAllFavouriteLinks";
import ViewAllDisscusionBoard from "./ViewAllDisscusionBoard";
import ViewAllEvents from "./ViewAllEvents";
import ViewAllAnnouncements from "./ViewAllAnnouncements";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

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
  Form,
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
    return new BpclDepartmentService(sp, props.context);
  }, [sp, props.context]);

  // const [navigationMenu, setNavigationMenu] = React.useState<INavigationMenu[]>(
  //   [],
  // );

  const [showAllMenus, setShowAllMenus] = React.useState(true);

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

  const [menus, setMenus] = React.useState<INavigationMenu[]>([]);

  const [selectedDiscussion, setSelectedDiscussion] =
    React.useState<IDiscussionBoard | null>(null);

  const [replies, setReplies] = React.useState<IDiscussionBoard[]>([]);

  // const [replyText, setReplyText] =
  // React.useState("");

  const [showDiscussionModal, setShowDiscussionModal] = React.useState(false);

  const [showAddDiscussion, setShowAddDiscussion] = React.useState(false);

  const [discussionTitle, setDiscussionTitle] = React.useState("");

  const [discussionMessage, setDiscussionMessage] = React.useState("");

  const [isQuestion, setIsQuestion] = React.useState(false);

  // const [visibleMenus, setVisibleMenus] = React.useState<INavigationMenu[]>([]);
  // const [overflowMenus, setOverflowMenus] = React.useState<INavigationMenu[]>([]);
  // const [showOverflowMenus, setShowOverflowMenus] = React.useState(false);

  // const menuContainerRef = React.useRef<HTMLDivElement>(null);

  const [editingReplyId, setEditingReplyId] = React.useState<number | null>(
    null,
  );
  const [replyText, setReplyText] = React.useState("");

  const [currentPage, setCurrentPage] = React.useState<
    "Home" | "FavouriteLinks" | "Events" | "DiscussionBoard" | "Announcements"
  >("Home");

  const siteTitle = props.context.pageContext.web.title;

  // code to show and hide the edit button
  let commandBarObserver: MutationObserver | undefined;

  const hideCommandBar = (): void => {
    const hide = (): void => {
      const commandBar = document.getElementById("spCommandBar");
      if (commandBar) {
        commandBar.style.display = "none";
      }
      // Hide Left Navigation
      const leftNav = document.getElementById("spLeftNav");
      if (leftNav) {
        leftNav.style.display = "none";
      }
      // Hide Left Navigation CommentsWrapper
      const commentBar = document.getElementById("CommentsWrapper");
      if (commentBar) {
        commentBar.style.display = "none";
      }

      // Hide Site Header
      const spHeader = document.getElementById("spSiteHeader");
      if (spHeader) {
        spHeader.style.display = "none";
      }

      // Hide App Bar
      const appBar = document.getElementById("sp-appBar");
      if (appBar) {
        appBar.style.display = "none";
      }

      // Hide Footer
      const footer = document.querySelector(
        '[data-automationid="SimpleFooter"]',
      ) as HTMLElement;

      if (footer) {
        footer.style.display = "none";
      }
    };

    hide();

    commandBarObserver = new MutationObserver(() => {
      hide();
    });

    commandBarObserver.observe(document.body, {
      childList: true,
      subtree: true,
    });
  };

  const showCommandBar = (): void => {
    const commandBar = document.getElementById("spCommandBar");
    if (commandBar) {
      commandBar.style.display = "";
    }

    // const leftNav = document.getElementById("spLeftNav");
    // if (leftNav) {
    //   leftNav.style.display = "";
    // }

    if (commandBarObserver) {
      commandBarObserver.disconnect();
      commandBarObserver = undefined;
    }
  };

  const checkCommandBarSetting = async (): Promise<void> => {
    try {
      const url =
        `${props.context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('ShowAndHideCommandBar')/items?$select=HideCommandBar&$top=1`;

      const response: SPHttpClientResponse =
        await props.context.spHttpClient.get(
          url,
          SPHttpClient.configurations.v1,
        );

      const data = await response.json();

      if (data.value.length > 0) {
        if (data.value[0].HideCommandBar) {
          hideCommandBar();
        } else {
          showCommandBar();
        }
      }
    } catch (error) {
      console.error("Error checking command bar setting", error);
    }
  };

  //   React.useEffect(() => {
  //   loadData().catch(console.error);
  //   checkCommandBarSetting().catch(console.error);

  //   return () => {
  //     if (commandBarObserver) {
  //       commandBarObserver.disconnect();
  //     }
  //   };
  // }, []);

  React.useEffect(() => {
    loadNavigationMenu();
    loadWelcomeBanners();
    loadVisionMission();
    loadFavouriteLinks();
    loadEvents();
    loadDocuments();
    loadAnnouncements();
    loadDiscussionBoard();

    checkCommandBarSetting().catch(console.error);

    return () => {
      if (commandBarObserver) {
        commandBarObserver.disconnect();
      }
    };
  }, []);

  const loadNavigationMenu = async (): Promise<void> => {
    try {
      const data = await service.getNavigationMenu();

      //        const menuData = await service.getNavigationMenu();
      // setMenus(menuData);
      setMenus(data);

      //setNavigationMenu(data);

      //console.log("Navigation Menu", data);
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
        return "bi bi-file-earmark-pdf";

      case "doc":
      case "docx":
        return "bi bi-file-earmark-text";

      case "xls":
      case "xlsx":
        return "bi bi-file-earmark-excel";

      case "ppt":
      case "pptx":
        return "bi bi-filetype-ppt";

      case "png":
      case "jpg":
      case "jpeg":
        return "bi bi-file-earmark-image";

      case "zip":
        return "bi bi-file-earmark-zip";

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
    console.log("Discussion Data:", data);

    setDiscussions(data);
  };

  const openDiscussion = async (discussion: IDiscussionBoard) => {
    console.log("Selected Discussion:", discussion);

    setSelectedDiscussion(discussion);

    const data = await service.getReplies(discussion.Id);

    setReplies(data);

    setShowDiscussionModal(true);
  };

  const handleShowDiscussion = () => {
    setShowAddDiscussion(true);
  };

  const handleCloseDiscussion = () => {
    setShowAddDiscussion(false);

    setDiscussionTitle("");

    setDiscussionMessage("");

    setIsQuestion(false);
  };

  const handleSaveDiscussion = async () => {
    if (!discussionTitle.trim()) {
      alert("Please enter Subject");

      return;
    }

    if (!discussionMessage.trim()) {
      alert("Please enter Message");

      return;
    }

    await service.addDiscussion(discussionTitle, discussionMessage, isQuestion);

    await loadDiscussionBoard();

    handleCloseDiscussion();
  };

  // const handleReply = async () => {

  //     if (!selectedDiscussion || !replyText.trim()) {
  //         return;
  //     }

  //     await service.addReply(
  //         selectedDiscussion.Id,
  //         replyText
  //     );

  //     const replyData = await service.getReplies(
  //         selectedDiscussion.Id
  //     );

  //     const discussionData = await service.getDiscussionBoard();

  //     setDiscussions(discussionData);

  //     const updatedDiscussion = discussionData.find(
  //         x => x.Id === selectedDiscussion.Id
  //     );

  //     if (updatedDiscussion) {
  //         setSelectedDiscussion(updatedDiscussion);
  //     }

  //     setReplies(replyData);

  //     setReplyText("");

  // };

  const openAnnouncement = React.useCallback((item: IAnnouncement): void => {
    if (!item.FileUrl?.trim()) {
      return;
    }

    window.open(item.FileUrl, "_blank", "noopener,noreferrer");
  }, []);

  const openEvent = React.useCallback((item: IEvent): void => {
    if (!item.FileUrl?.trim()) {
      return;
    }

    window.open(item.FileUrl, "_blank", "noopener,noreferrer");
  }, []);

  const handleReply = async () => {
    if (!selectedDiscussion || !replyText.trim()) {
      return;
    }

    if (editingReplyId) {
      await service.updateReply(editingReplyId, replyText);
    } else {
      await service.addReply(selectedDiscussion.Id, replyText);
    }

    setEditingReplyId(null);

    setReplyText("");

    const replyData = await service.getReplies(selectedDiscussion.Id);

    setReplies(replyData);

    await loadDiscussionBoard();
  };

  const currentUserId = props.context.pageContext.legacyPageContext.userId;

  //   if (currentPage === "FavouriteLinks") {

  //     return (
  //         <ViewAllFavouriteLinks
  //             context={props.context}
  //             onBack={() => setCurrentPage("Home")}
  //         />
  //     );

  // }

  if (currentPage === "DiscussionBoard") {
    return (
      <ViewAllDisscusionBoard
        context={props.context}
        onBack={() => setCurrentPage("Home")}
      />
    );
  }

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
      {/* <Navbar expand="lg" className={styles.navbarCustom}>
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
      </Navbar> */}

      
      {/* <div
  className={`${styles.navigationWrapper} ${
    showAllMenus ? styles.expanded : ""
  }`}
> */}
  <div
        className={`${styles.menuRows} ${
            showAllMenus ? styles.expanded : ""
        }`}
    >
  {/* <Nav className={styles.navWrapper}> */}
    {/* Existing menu rendering */}
    <Navbar expand="lg" className={`${styles.navbarCustom}`}>
        <Container fluid>
          <Navbar.Brand className={styles.brandText}>{siteTitle}</Navbar.Brand>

          <Navbar.Toggle aria-controls="basic-navbar-nav" />

          <Navbar.Collapse id="basic-navbar-nav">
            <Nav className={`ms-auto ${styles.navWrapper}`}>
              {(() => {
                const parents = menus.filter((menu) => !menu.ParentIDId);

                return parents.map((parent, index) => {
                  const children = menus.filter(
                    (x) => x.ParentIDId === parent.Id,
                  );

                  if (children.length > 0) {
                    return (
                      <NavDropdown
                        key={parent.Id}
                        title={parent.Title}
                        id={`menu-${parent.Id}`}
                        className={
                          index === parents.length - 1
                            ? styles.lastDropdown
                            : ""
                        }
                      >
                        {children.map((child) => (
                          <NavDropdown.Item
                            key={child.Id}
                            href="#"
                            onClick={(e) => {
                              e.preventDefault();

                              if (child.RedirectURL?.Url) {
                                window.open(
                                  child.RedirectURL.Url,
                                  "_blank",
                                  "noopener,noreferrer",
                                );
                              }
                            }}
                          >
                            {child.Title}
                          </NavDropdown.Item>
                        ))}
                      </NavDropdown>
                    );
                  }

                  return (
                    <Nav.Link
                      key={parent.Id}
                      href="#"
                      onClick={(e) => {
                        e.preventDefault();

                        if (parent.RedirectURL?.Url) {
                          window.open(
                            parent.RedirectURL.Url,
                            "_blank",
                            "noopener,noreferrer",
                          );
                        }
                      }}
                    >
                      {parent.Title}
                    </Nav.Link>
                  );
                });
              })()}
            </Nav>
          </Navbar.Collapse>
        </Container>
      </Navbar>

  {/* </Nav> */}

  {menus.filter(m => !m.ParentIDId).length > 8 && (
    <button
      className={styles.expandButton}
      onClick={() => setShowAllMenus(!showAllMenus)}
    >
      <i
        className={`bi ${
          showAllMenus ? "bi-chevron-up" : "bi-chevron-down"
        }`}
      ></i>
    </button>
  )}
</div>



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
            <span className={`${styles.portalText} mx-2 `}>Portal</span>
          </h1>

          <h3 className={styles.heroSubtitle}>
            Your one-stop platform for updates, resources and collaboration
          </h3>

          <p className={styles.heroDescription}>
            {activeSlide?.WelcomeBannerDesc}
          </p>

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
      <Container fluid className="p-4 px-lg-5">
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
                      textAlign: "center",
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
                          <span className={styles.linkDesc}>{link.Title}</span>
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
      <Container fluid className="p-4 px-lg-5 ">
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
                  const sameDay = start.toDateString() === end.toDateString();

                  return (
                    <Card key={event.Id} className={styles.eventCard}>
                      <div
                        className={styles.eventCard}
                        onClick={() => openEvent(event)}
                        style={{
                          cursor: event.FileUrl ? "pointer" : "default",
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
                                  {sameDay
                                    ? start
                                        .getDate()
                                        .toString()
                                        .padStart(2, "0")
                                    : `${start.getDate().toString().padStart(2, "0")} - ${end
                                        .getDate()
                                        .toString()
                                        .padStart(2, "0")}`}
                                </div>

                                <div className={styles.month}>
                                  {sameDay
                                    ? start.toLocaleString("default", {
                                        month: "short",
                                      })
                                    : `${start.toLocaleString("default", {
                                        month: "short",
                                      })} - ${end.toLocaleString("default", {
                                        month: "short",
                                      })}`}
                                </div>
                              </div>
                            </Col>

                            <Col>
                              <h2
                                className={styles.eventTitle}
                                title={event.Title}
                              >
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
                      </div>
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
              </div>

              <div className="flex-grow-1 ">
                {documents.length === 0 ? (
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
                      style={{
                        cursor: doc.FileUrl?.trim() ? "pointer" : "default",
                      }}
                      onClick={() => {
                        if (!doc.FileUrl?.trim()) return;
                        window.open(
                          doc.FileUrl,
                          "_blank",
                          "noopener,noreferrer",
                        );
                      }}
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
                          href={doc.FileUrl || "#"}
                          onClick={(e) => {
                            e.preventDefault();

                            if (!doc.FileUrl?.trim()) return;

                            window.open(
                              doc.FileUrl,
                              "_blank",
                              "noopener,noreferrer",
                            );
                          }}
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
      <Container fluid className="p-4 px-lg-5">
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
                        <div
                          className={styles.newsCard}
                          onClick={() => openAnnouncement(item)}
                          style={{
                            cursor: item.FileUrl ? "pointer" : "default",
                          }}
                        >
                          <div className={styles.cardImageWrapper}>
                            <img
                              src={item.ImageUrl || news1}
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
                        </div>
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
                    href="#"
                    className={styles.viewAllLink}
                    onClick={(e) => {
                      e.preventDefault();
                      handleShowDiscussion();
                    }}
                  >
                    <i className="bi bi-plus-circle me-2"></i>
                    Add
                  </a>
                  <a
                    href="#"
                    className={styles.viewAllLink}
                    onClick={(e) => {
                      e.preventDefault();
                      setCurrentPage("DiscussionBoard");
                    }}
                  >
                    View All
                  </a>
                </div>

                {discussions.slice(0, 3).map((item) => (
                  <div
                    key={item.Id}
                    className={styles.discussionCard}
                    style={{ cursor: "pointer" }}
                    onClick={() => openDiscussion(item)}
                  >
                    <div className={styles.statusBar} />

                    <div className={styles.content}>
                      <div className="row align-items-center">
                        <div className="col-12">
                          <h3 className={styles.title} title={item.Title}>
                            {item.Title}
                          </h3>

                          <div className={styles.metadata}>
                            Created By :
                            <span className={styles.author}>
                              {item.Author?.Title}
                            </span>
                            <span className={styles.separator}>•</span>
                            Reply Count :
                            <span className={styles.author}>
                              {item.ReplyCount}
                            </span>
                            {/* <span className={styles.separator}>•</span> */}
                            <span className={styles.date}>
                              {item.Created
                                ? new Date(item.Created).toLocaleDateString(
                                    "en-GB",
                                  )
                                : ""}
                            </span>
                            <span className={styles.time}>
                              {item.Created
                                ? new Date(item.Created).toLocaleTimeString(
                                    [],
                                    {
                                      hour: "2-digit",
                                      minute: "2-digit",
                                    },
                                  )
                                : ""}
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
          <Modal
            show={showAddDiscussion}
            onHide={handleCloseDiscussion}
            centered
            size="xl"
          >
            <Modal.Header closeButton>
              <Modal.Title>Add New Discussion</Modal.Title>
            </Modal.Header>

            <Modal.Body>
              <Form>
                {/* Subject */}
                <Form.Group className="mb-3">
                  <Form.Label>Subject</Form.Label>
                  <Form.Control
                    type="text"
                    placeholder="Enter subject"
                    value={discussionTitle}
                    onChange={(e) => setDiscussionTitle(e.target.value)}
                  />{" "}
                </Form.Group>

                {/* Body */}
                <Form.Group className="mb-3">
                  <Form.Label>Body</Form.Label>
                  <Form.Control
                    as="textarea"
                    rows={5}
                    placeholder="Enter your message..."
                    value={discussionMessage}
                    onChange={(e) => setDiscussionMessage(e.target.value)}
                  />
                </Form.Group>

                {/* Question Checkbox */}
                <Form.Group className="mb-2">
                  <Form.Check
                    type="checkbox"
                    id="question-checkbox"
                    label="Question"
                    checked={isQuestion}
                    onChange={(e) => setIsQuestion(e.target.checked)}
                    className={styles.questionCheckbox}
                  />
                </Form.Group>
              </Form>
            </Modal.Body>

            <Modal.Footer>
              <Button
                className={styles.cancelBtn}
                onClick={handleCloseDiscussion}
              >
                Cancel
              </Button>

              <Button className={styles.saveBtn} onClick={handleSaveDiscussion}>
                Save
              </Button>
            </Modal.Footer>
          </Modal>
        </Row>

        <Modal
          show={showDiscussionModal}
          onHide={() => {
            setReplyText("");
            setEditingReplyId(null);
            setShowDiscussionModal(false);
          }}
          centered
          size="lg"
        >
          <Modal.Header closeButton>
            <Modal.Title className={styles.discussionTitle}>
              {selectedDiscussion?.Title}
            </Modal.Title>
          </Modal.Header>

          <Modal.Body
            className="d-flex flex-column"
            style={{ maxHeight: "80vh" }}
          >
            {/* Discussion */}
            <div className="d-flex align-items-start mb-4">
              <img
                src="https://www.w3schools.com/howto/img_avatar.png"
                alt="User"
                className="rounded-circle"
                width={50}
                height={50}
              />

              <div className="ms-3 flex-grow-1">
                <div className="d-flex justify-content-between align-items-center">
                  <h6 className={`${styles.userName} mb-1`}>
                    {selectedDiscussion?.Author?.Title ?? "Unknown User"}
                  </h6>

                  <small className="text-muted">
                    {selectedDiscussion?.Created
                      ? new Date(selectedDiscussion.Created).toLocaleDateString(
                          "en-GB",
                        ) +
                        " • " +
                        new Date(selectedDiscussion.Created).toLocaleTimeString(
                          [],
                          {
                            hour: "2-digit",
                            minute: "2-digit",
                          },
                        )
                      : ""}
                  </small>
                </div>
                <p className={`${styles.replyText} mb-0`}>
                  {selectedDiscussion?.Message}
                </p>
              </div>
            </div>

            <h6 className="mb-3">All Replies</h6>

            {/* Scrollable Replies */}
            {/* <div
                className="flex-grow-1 overflow-auto pe-2"
                style={{ maxHeight: "350px" }}
              > */}

            {/* Reply 1 */}
            {/* <div className="d-flex align-items-start mb-4">
                 {replies.map((reply) => (

<div
    key={reply.Id}
    className="d-flex align-items-start mb-4"
>

    <img
        src="https://www.w3schools.com/howto/img_avatar.png"
        alt=""
        className="rounded-circle"
        width={40}
        height={40}
    />

    <div className="ms-3 flex-grow-1">

        <h6 className={`${styles.userName} mb-1`}>
            {reply.Author?.Title}
        </h6>

        <p className={`${styles.replyText} mb-2`}>
            {reply.Message}
        </p>

        <div className="d-flex align-items-center gap-2 small text-muted">

            <span>
                {reply.Created
                    ? new Date(reply.Created).toLocaleDateString("en-GB")
                    : ""}
            </span>

            <span>•</span>

            <span>
                {reply.Created
                    ? new Date(reply.Created).toLocaleTimeString([], {
                          hour: "2-digit",
                          minute: "2-digit",
                      })
                    : ""}
            </span>

        </div>

    </div>

</div>

))}
                </div> */}
            <div className={styles.replyContainer}>
              {replies.length === 0 && (
                <div className="text-center text-muted py-4">
                  No replies available.
                </div>
              )}

              {replies.map((reply) => (
                <div key={reply.Id} className={styles.replyItem}>
                  <img
                    src="https://www.w3schools.com/howto/img_avatar.png"
                    className={styles.replyAvatar}
                    alt=""
                  />

                  <div className={styles.replyContent}>
                    <div className="d-flex justify-content-between">
                      <h6 className={styles.userName}>{reply.Author?.Title}</h6>
                    </div>

                    <div className={styles.replyText}>{reply.Message}</div>

                    <div className={styles.replyMeta}>
                      <span>
                        {reply.Created
                          ? new Date(reply.Created).toLocaleDateString("en-GB")
                          : ""}
                      </span>

                      <span>•</span>

                      <span>
                        {reply.Created
                          ? new Date(reply.Created).toLocaleTimeString([], {
                              hour: "2-digit",
                              minute: "2-digit",
                            })
                          : ""}
                      </span>

                      <span>•</span>

                      {/* <a href="#">Reply</a> */}
                      <Button
                        variant="link"
                        size="sm"
                        className="p-0 text-decoration-none"
                        onClick={() => {
                          setEditingReplyId(null); // Ensure we're adding a new reply, not editing
                          setReplyText(`@${reply.Author?.Title} `);
                        }}
                      >
                        Reply
                      </Button>

                      <span>•</span>

                      {/* <a href="#">Edit</a> */}
                      {reply.Author?.Id === currentUserId && (
                        <Button
                          variant="link"
                          size="sm"
                          className="p-0 text-decoration-none"
                          onClick={() => {
                            setEditingReplyId(reply.Id);

                            setReplyText(reply.Message);
                          }}
                        >
                          Edit
                        </Button>
                      )}
                    </div>
                  </div>
                </div>
              ))}
            </div>
            {/* Reply 2 */}
            {/* <div className="d-flex align-items-start mb-4">
                  <img
                    src="https://www.w3schools.com/howto/img_avatar.png"
                    alt="User"
                    className="rounded-circle"
                    width={40}
                    height={40}
                  />
 
                  <div className="ms-3 flex-grow-1">
                    <h6 className={` ${styles.userName} mb-1 `}>
                      Jeyaprakash G
                    </h6>
 
                    <p className={`${styles.replyText} mb-2`}>
                      Lorem ipsum dolor sit amet consectetur adipiscing elit.
                      Sit amet consectetur adipiscing elit quisque faucibus ex.
                      Adipiscing elit quisque faucibus ex sapien vitae
                      pellentesque.
                    </p>
 
                    <div className="d-flex align-items-center gap-2 small text-muted">
                      <span>15 Jul 2025</span>
                      <span>•</span>
                      <span>11:20 AM</span>
                      <span>•</span>
 
                      <Button
                        variant="link"
                        size="sm"
                        className="p-0 text-muted text-decoration-none"
                      >
                        Reply
                      </Button>
 
                      <span>•</span>
 
                      <Button
                        variant="link"
                        size="sm"
                        className="p-0 text-muted text-decoration-none"
                      >
                        Edit
                      </Button>
                    </div>
                  </div>
                </div> */}

            {/* Reply 3 */}
            {/* <div className="d-flex align-items-start mb-4">
                  <img
                    src="https://www.w3schools.com/howto/img_avatar.png"
                    alt="User"
                    className="rounded-circle"
                    width={40}
                    height={40}
                  />
 
                  <div className="ms-3 flex-grow-1">
                    <h6 className={` ${styles.userName} mb-1 `}>
                      Aravindhan S
                    </h6>
 
                    <p className={`${styles.replyText} mb-2`}>
                      Lorem ipsum dolor sit amet consectetur adipiscing elit.
                      Sit amet consectetur adipiscing elit quisque faucibus ex.
                      Adipiscing elit quisque faucibus ex sapien vitae
                      pellentesque.
                    </p>
 
                    <div className="d-flex align-items-center gap-2 small text-muted">
                      <span>15 Jul 2025</span>
                      <span>•</span>
                      <span>11:35 AM</span>
                      <span>•</span>
 
                      <Button
                        variant="link"
                        size="sm"
                        className="p-0 text-muted text-decoration-none"
                      >
                        Reply
                      </Button>
 
                      <span>•</span>
 
                      <Button
                        variant="link"
                        size="sm"
                        className="p-0 text-muted text-decoration-none"
                      >
                        Edit
                      </Button>
                    </div>
                  </div>
                </div> */}

            {/* Reply 4 */}
            {/* <div className="d-flex align-items-start mb-4">
                  <img
                    src="https://www.w3schools.com/howto/img_avatar.png"
                    alt="User"
                    className="rounded-circle"
                    width={40}
                    height={40}
                  />
 
                  <div className="ms-3 flex-grow-1">
                    <h6 className={` ${styles.userName} mb-1 `}>
                      Silpa Chandrasekaran
                    </h6>
 
                    <p className={`${styles.replyText} mb-2`}>
                      Lorem ipsum dolor sit amet consectetur adipiscing elit.
                      Sit amet consectetur adipiscing elit quisque faucibus ex.
                      Adipiscing elit quisque faucibus ex sapien vitae
                      pellentesque.
                    </p>
 
                    <div className="d-flex align-items-center gap-2 small text-muted">
                      <span>15 Jul 2025</span>
                      <span>•</span>
                      <span>11:50 AM</span>
                      <span>•</span>
 
                      <Button
                        variant="link"
                        size="sm"
                        className="p-0 text-muted text-decoration-none"
                      >
                        Reply
                      </Button>
 
                      <span>•</span>
 
                      <Button
                        variant="link"
                        size="sm"
                        className="p-0 text-muted text-decoration-none"
                      >
                        Edit
                      </Button>
                    </div>
                  </div>
                </div> */}

            {/* Reply 5 */}
            {/* <div className="d-flex align-items-start mb-4">
                  <img
                    src="https://www.w3schools.com/howto/img_avatar.png"
                    alt="User"
                    className="rounded-circle"
                    width={40}
                    height={40}
                  />
 
                  <div className="ms-3 flex-grow-1">
                    <h6 className={` ${styles.userName} mb-1 `}>Pournami</h6>
 
                    <p className={`${styles.replyText} mb-2`}>
                      Lorem ipsum dolor sit amet consectetur adipiscing elit.
                      Sit amet consectetur adipiscing elit quisque faucibus ex.
                      Adipiscing elit quisque faucibus ex sapien vitae
                      pellentesque.
                    </p>
 
                    <div className="d-flex align-items-center gap-2 small text-muted">
                      <span>15 Jul 2025</span>
                      <span>•</span>
                      <span>11:50 AM</span>
                      <span>•</span>
 
                      <Button
                        variant="link"
                        size="sm"
                        className="p-0 text-muted text-decoration-none"
                      >
                        Reply
                      </Button>
 
                      <span>•</span>
 
                      <Button
                        variant="link"
                        size="sm"
                        className="p-0 text-muted text-decoration-none"
                      >
                        Edit
                      </Button>
                    </div>
                  </div>
                </div> */}

            {/* Reply Box */}
            <Form.Group className="mt-3">
              <Form.Label>Reply</Form.Label>

              {/* <Form.Control
                  as="textarea"
                  rows={4}
                  placeholder="Write your reply..."
                /> */}
              <Form.Control
                as="textarea"
                rows={4}
                placeholder="Write your reply..."
                value={replyText}
                onChange={(e) => setReplyText(e.target.value)}
              />
            </Form.Group>

            {/* Buttons */}
            <div className="d-flex justify-content-end gap-2 mt-3">
              <Button
                variant="secondary"
                onClick={() => {
                  setReplyText("");
                  setEditingReplyId(null);
                  setShowDiscussionModal(false);
                }}
                className={styles.cancelBtn}
              >
                Cancel
              </Button>

              {/* <Button className={styles.saveBtn}>Reply</Button> */}
              <Button
                className={styles.saveBtn}
                onClick={handleReply}
                disabled={!replyText.trim()}
              >
                Reply
              </Button>
            </div>
          </Modal.Body>
        </Modal>
      </Container>

      {/* Organogram Modal  */}
      <Modal show={show} onHide={handleClose} centered size="xl">
        <Modal.Header closeButton>
          <Modal.Title className={styles.modalTitle}>
            Organogram of {siteTitle}
          </Modal.Title>
        </Modal.Header>

        <Modal.Body className={`px-1 ${styles.modalBody}`}>
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
            height="500px"
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
