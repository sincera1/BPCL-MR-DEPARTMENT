import * as React from "react";
import styles from "./BpclMrDepartment.module.scss";
//import styles from "./ViewAllPagesCommon.module.scss";

//import { useState } from "react";

import "@fontsource/inter/500.css";
import "@fontsource/inter/600.css";
import "@fontsource/inter/700.css";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap-icons/font/bootstrap-icons.css";

import { Row, Col, Button, Pagination, Modal, Form } from "react-bootstrap";

import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp";

import BpclDepartmentService, {
  IDiscussionBoard,
} from "../services/BpclDepartmentService";

import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IViewAllDisscusionBoardProps {
  context: WebPartContext;
  onBack: () => void;
}

const ViewAllDisscusionBoard: React.FC<IViewAllDisscusionBoardProps> = (
  props,
) => {
  const sp: SPFI = React.useMemo(() => {
    return spfi().using(SPFx(props.context));
  }, [props.context]);

  const service = React.useMemo(() => {
    return new BpclDepartmentService(sp, props.context);
  }, [sp, props.context]);

  const [discussions, setDiscussions] = React.useState<IDiscussionBoard[]>([]);

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

  const [currentPage, setCurrentPage] = React.useState(1);

  const [pageSize, setPageSize] = React.useState(10);

  const [editingReplyId, setEditingReplyId] = React.useState<number | null>(
    null,
  );
  const [replyText, setReplyText] = React.useState("");

  React.useEffect(() => {
    loadDiscussionBoard().catch(console.error);
  }, []);


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
      // alert("Please enter Subject");

      return;
    }

    if (!discussionMessage.trim()) {
      // alert("Please enter Message");

      return;
    }

    await service.addDiscussion(discussionTitle, discussionMessage, isQuestion);

    await loadDiscussionBoard();

    handleCloseDiscussion();
  };

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
    const latest = await service.getDiscussionBoard();

    const current = latest.find((x) => x.Id === selectedDiscussion.Id);

    if (current) {
      setSelectedDiscussion(current);
    }
  };

  const currentUserId = props.context.pageContext.legacyPageContext.userId;

  

  const totalPages = Math.ceil(discussions.length / pageSize);

  const pagedDiscussions = discussions.slice(
    (currentPage - 1) * pageSize,
    currentPage * pageSize,
  );

  return (
    <div className={styles.pageContainer}>
      <div className={styles.banner}>
        <button className={styles.backIcon} onClick={props.onBack}>
          <i className="bi bi-chevron-left"></i>
        </button>

        <div className={styles.bannerIcon}>
          <i className="bi bi-chat-dots-fill"></i>
        </div>

        <div>
          <h3>Discussion Board</h3>
          <p>Participate in discussions and collaborate with team members.</p>
        </div>
      </div>

      <div className={styles.cardContainer}>
        {pagedDiscussions.length > 0 && (
          <Row>
            <Col xs={12}>
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
                </div>

                {pagedDiscussions.map((item) => (
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
          </Row>
        )}

        {pagedDiscussions.length === 0 && (
          <div
            className="text-center py-5"
            style={{
              color: "#6c757d",
              fontWeight: 500,
            }}
          >
            No Discussions Available
          </div>
        )}

{/*  add modal */}
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

{/* view modal */}
        <Modal
          show={showDiscussionModal}
          onHide={() => {
           
            setReplyText("");
            setEditingReplyId(null);
            setSelectedDiscussion(null);
            setReplies([]);
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
              {/* <img
                src="https://www.w3schools.com/howto/img_avatar.png"
                alt="User"
                className="rounded-circle"
                width={50}
                height={50}
              /> */}
              <img
  src={`${props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${selectedDiscussion?.Author?.EMail}`}
  alt="User"
  className="rounded-circle"
  width={50}
  height={50}
  style={{ objectFit: "cover" }}
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

            <div className={styles.replyContainer}>
              {replies.length === 0 && (
                <div className="text-center text-muted py-4">
                  No replies available.
                </div>
              )}

              {replies.map((reply) => (
                <div key={reply.Id} className={styles.replyItem}>
                  {/* <img
                    src="https://www.w3schools.com/howto/img_avatar.png"
                    className={styles.replyAvatar}
                    alt=""
                  /> */}
                  <img
  src={`${props.context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${reply.Author?.EMail}`}
  className={styles.replyAvatar}
  alt=""
  style={{ objectFit: "cover" }}
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

            {/* Reply Box */}
            <Form.Group className="mt-3">
              <Form.Label>Reply</Form.Label>

             
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
                  setSelectedDiscussion(null);
                  setReplies([]);
                  setShowDiscussionModal(false);
                }}
                className={styles.cancelBtn}
              >
                Cancel
              </Button>

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
              disabled={currentPage === 1}
              onClick={() => setCurrentPage(1)}
            />

            <Pagination.Prev
              disabled={currentPage === 1}
              onClick={() => setCurrentPage(currentPage - 1)}
            />

            {Array.from({ length: totalPages }, (_, index) => (
              <Pagination.Item
                key={index}
                active={currentPage === index + 1}
                onClick={() => setCurrentPage(index + 1)}
              >
                {index + 1}
              </Pagination.Item>
            ))}
            <Pagination.Next
              disabled={currentPage === totalPages}
              onClick={() => setCurrentPage(currentPage + 1)}
            />

            <Pagination.Last
              disabled={currentPage === totalPages}
              onClick={() => setCurrentPage(totalPages)}
            />
          </Pagination>
        </div>
      </div>
    </div>
  );
};

export default ViewAllDisscusionBoard;
