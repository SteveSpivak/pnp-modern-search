import * as React from "react";
import styles from "./LinkedinPostCard.module.scss";

export interface ILinkedinPostCardProps {
  authorName: string;
  authorTitle: string;
  postText: string;
  postUrl: string;
  imageUrl: string;
  date: string;
}

export const LinkedinPostCard: React.FC<ILinkedinPostCardProps> = (props) => {
  return (
    <div className={styles.linkedinCard}>
      <div className={styles.header}>
        <div className={styles.authorAvatar}>
          {/* Default avatar initials */}
          {props.authorName.charAt(0)}
        </div>
        <div className={styles.authorInfo}>
          <div className={styles.authorName}>{props.authorName}</div>
          {props.authorTitle && <div className={styles.authorTitle}>{props.authorTitle}</div>}
          {props.date && <div className={styles.postDate}>{props.date}</div>}
        </div>
      </div>

      <div className={styles.postBody}>
        {props.postText}
      </div>

      {props.imageUrl && (
        <div className={styles.postImage}>
          <img src={props.imageUrl} alt="Post content" />
        </div>
      )}

      <div className={styles.footer}>
        <a href={props.postUrl} target="_blank" rel="noopener noreferrer">
          View on LinkedIn
        </a>
      </div>
    </div>
  );
};
