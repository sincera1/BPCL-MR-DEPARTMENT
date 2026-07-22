import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";

import { graphfi } from "@pnp/graph";
import { SPFx } from "@pnp/graph";
import "@pnp/graph/taxonomy";
import { WebPartContext } from "@microsoft/sp-webpart-base";

/* ============================
   Navigation Menu Interface
============================ */

export interface INavigationMenu { 
  Id: number;
  Title: string;
  RedirectURL: {
    Url: string;
    Description: string;
  };
  Sequence: number;
  ParentIDId?: number;
  IsActive: boolean;
}

export interface ITermNavigation {
  Id: string;
  Title: string;
  ParentId?: string;
  Children?: ITermNavigation[];
}

/* ============================
   Welcome Banner Interface
============================ */

export interface IWelcomeBanner {
    Id: number;
    WelcomeBannerTitle: string;
    WelcomeBannerDesc: string;
    ImageUrl: string;
    Order: number;
    IsActive: boolean;
}

/* ============================
   Mission Vision Interface
============================ */

export interface IVissionMission {
  Id: number;
  Title: string;
  Description: string;
  IsActive: boolean;
  AttachmentUrl?: string;
}

export interface IAnnouncementAttachment {
  FileName: string;
  ServerRelativeUrl: string;
}

export interface IAnnouncement {

  Id: number;

  Title: string;

  Description: string;

  PublishedDate?: string;

  ImageUrl?: string;

  FileUrl?: string;

  IsActive: boolean;

  Attachments: IAnnouncementAttachment[];

}

export interface IEventAttachment {
  FileName: string;
  ServerRelativeUrl: string;
}

export interface IEvent {

  Id: number;

  Title: string;

  Description: string;


  StartDate: string;

  EndDate?: string;

  IsActive: boolean;

  FileUrl?: string;

  Attachments: IEventAttachment[];

}

export interface IFavouriteLink {

    Id:number;

    Title:string;

    RedirectURL:{
        Url:string;
        Description:string;
    };

    DisplayOrder:number;

    IsActive:boolean;

}

export interface ISharedDocument {

  Id: number;

  LibraryName: string;

  FileName: string;

  FileUrl: string;

  Modified: string;

  FileType: string;

}

export interface IDiscussionBoard {

    Id: number;

    Title: string;

    Message: string;

    ParentId?: number;

    IsReply: boolean;

    ReplyCount: number;

     DisplayOrder: number;

    IsActive: boolean;

    Author?: {
  Id: number;
  Title: string;
  EMail: string;
};

Created: string;
}

/* ============================
   Service Class
============================ */

export default class BpclDepartmentService {


// constructor(
//     private sp: SPFI,
  
// ) {}

private context: WebPartContext;

constructor(
    private sp: SPFI,
    context: WebPartContext
) {
    this.context = context;
}

  /* ============================
     Navigation Menu
  ============================ */ 

  public async getNavigationMenu(): Promise<INavigationMenu[]> {

    try {

      const items = await this.sp.web.lists
        .getByTitle("MR_SL_NavigationMenu")
        .items
        .select(
          "Id",
          "Title",
          "RedirectURL",
          "Sequence",
          "ParentIDId",
          "IsActive"
        )
        .filter("IsActive eq 1")
        .orderBy("Sequence", true)();

      return items as INavigationMenu[];

    } catch (error) {

      console.error(
        "Error fetching Navigation Menu",
        error
      );

      return [];
    }
  }

/* ============================
   Term Store Navigation
============================ */

public async getTermNavigation(
    termGroupName: string,
    termSetName: string
): Promise<INavigationMenu[]> {

    try {

        const graph = graphfi().using(SPFx(this.context));

        // Get all groups
        const groups = await graph.termStore.groups();

       const group = groups.find((g: any) =>
    g.displayName === termGroupName ||
    g.name === termGroupName
);

if (!group || !group.id) {
    console.log("Group not found");
    return [];
}

const groupId = group.id as string;

const sets = await graph.termStore
    .groups
    .getById(groupId)
    .sets();

        const set = sets.find((s: any) =>
            s.localizedNames?.[0]?.name === termSetName ||
            s.name === termSetName
        );

        if (!set) {
            console.log("Term Set not found");
            return [];
        }

       

        if (!set || !set.id) {
    console.log("Term Set not found");
    return [];
}

const setId = set.id as string;

const tree = await graph.termStore
    .groups
    .getById(groupId)
    .sets
    .getById(setId)
    .getAllChildrenAsTree();

        const menus: INavigationMenu[] = [];

        const build = (
            items: any[],
            parentId?: any
        ) => {

            items.forEach((term: any) => {

                let redirect = "";

                if (term.properties) {

                    const prop = term.properties.find(
                        (p: any) =>
                            p.key === "redirectUrl" ||
                            p.key === "RedirectUrl"
                    );

                    if (prop) {
                        redirect = prop.value;
                    }
                }

                menus.push({

                    Id: term.id,

                    Title:
                        term.labels?.find((x: any) => x.isDefault)?.name ||
                        term.labels?.[0]?.name ||
                        "",

                    ParentIDId: parentId,

                    RedirectURL: {
                        Url: redirect,
                        Description: ""
                    },

                    Sequence: 0,

                    IsActive: true

                });

                if (
                    term.children &&
                    term.children.length > 0
                ) {

                    build(
                        term.children,
                        term.id
                    );

                }

            });

        };

        build(tree);

        console.log("Term Navigation", menus);

        return menus;

    }
    catch (e) {

        console.log(e);

        return [];

    }
 
}
  /* ============================
     Welcome Banner
  ============================ */

public async getWelcomeBanners(): Promise<IWelcomeBanner[]> {

    try {

        const items = await this.sp.web.lists
            .getByTitle("MR_DL_WelcomeBanner")
            .items
            .select(
                "Id",
                "WelcomeBannerTitle",
                "WelcomeBannerDesc",
                "Order",
                "IsActive",
                "FileRef"
            )
            .filter("IsActive eq 1")
            .orderBy("Order", true)();

        return items.map((item: any) => ({

            Id: item.Id,
            WelcomeBannerTitle: item.WelcomeBannerTitle,
            WelcomeBannerDesc: item.WelcomeBannerDesc,
            Order: item.Order,
            IsActive: item.IsActive,
            ImageUrl: item.FileRef

        }));

    }
    catch (e) {

        console.log(e);
        return [];

    }

}

  /* ============================
     Mission & Vision
  ============================ */

  public async getVisionMission(): Promise<IVissionMission[]> {

    try {

      const items = await this.sp.web.lists
        .getByTitle("MR_SL_VissionMission")
        .items
        .select(
          "Id",
          "Title",
          "Description",
          "IsActive",
          "AttachmentFiles"
        )
        .expand("AttachmentFiles")
        .filter("IsActive eq 1")();

      return items.map((item: any) => ({

        Id: item.Id,

        Title: item.Title,

        Description: item.Description,

        IsActive: item.IsActive,

       AttachmentUrl:
  item.AttachmentFiles &&
  item.AttachmentFiles.length > 0
    ? `${window.location.origin}${item.AttachmentFiles[0].ServerRelativeUrl}`
    : ""

      }));

    } catch (error) {

      console.error(
        "Error fetching Mission Vision",
        error
      );

      return [];
    }
  }
    /* ============================
      Announcements
  ============================ */ 
// this is for home page
 public async getAnnouncements(): Promise<IAnnouncement[]> {

  try {

    const items = await this.sp.web.lists
      .getByTitle("MR_SL_Announcement")
      .items
      .select(
        "Id",
        "Title",
        "Description",
        "PublishedDate",
        "IsActive",
        "AttachmentFiles"
      )
      .expand("AttachmentFiles")
      .filter("IsActive eq 1")
      .orderBy("PublishedDate", false)();

  return items.map((item: any) => {

  const attachments = item.AttachmentFiles || [];

  let imageUrl = "";
  let fileUrl = "";

  attachments.forEach((file: any) => {

    const extension = file.FileName
      .split(".")
      .pop()
      ?.toLowerCase(); 

    if (
      !imageUrl &&
      ["jpg", "jpeg", "png", "gif", "webp"].includes(extension || "")
    ) {
      imageUrl =
        `${window.location.origin}${file.ServerRelativeUrl}`;
    }

    if (
      !fileUrl &&
      ["pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx"]
        .includes(extension || "")
    ) {
      fileUrl =
        `${window.location.origin}${file.ServerRelativeUrl}`;
    }


  });
   // If no document was found, open the image itself
if (!fileUrl) {
  fileUrl = imageUrl;
}

  return {

    Id: item.Id,

    Title: item.Title,

    Description: item.Description,

    PublishedDate: item.PublishedDate,

    IsActive: item.IsActive,

    ImageUrl: imageUrl,

    FileUrl: fileUrl,

    Attachments: attachments.map((file: any) => ({

      FileName: file.FileName,

      ServerRelativeUrl: file.ServerRelativeUrl

    }))

  };

});

  } catch (error) {

    console.error("Error fetching announcements", error);

    return [];
  }
}
 
// this is for view all
public async getAllAnnouncements(): Promise<IAnnouncement[]> {

  try {

    const items = await this.sp.web.lists
      .getByTitle("MR_SL_Announcement")
      .items
      .select(
        "Id",
        "Title",
        "Description",
        "PublishedDate",
        "IsActive",
        "AttachmentFiles"
      )
      .expand("AttachmentFiles")
      .filter("IsActive eq 1")
      .orderBy("PublishedDate", false)();

    return items.map((item: any) => {

      let imageUrl = "";
      let fileUrl = "";

      const attachments = item.AttachmentFiles || [];

      attachments.forEach((file: any) => {

        const ext = file.FileName.split(".").pop()?.toLowerCase();

        const url = `${window.location.origin}${file.ServerRelativeUrl}`;

        if (
          ["jpg", "jpeg", "png", "gif", "bmp", "webp", "svg"]
            .includes(ext || "")
        ) {

          if (!imageUrl) {
            imageUrl = url;
          }

        } else {

          if (!fileUrl) {
            fileUrl = url;
          }

        }

      });

      if (!fileUrl) {
        fileUrl = imageUrl;
      }

      return {

        Id: item.Id,
        Title: item.Title,
        Description: item.Description,
        PublishedDate: item.PublishedDate,
        ImageUrl: imageUrl,
        FileUrl: fileUrl,
        IsActive: item.IsActive,
        Attachments: attachments.map((file: any) => ({
          FileName: file.FileName,
          ServerRelativeUrl: file.ServerRelativeUrl
        }))

      };

    });

  }

  catch (error) {

    console.log(error);
    return [];

  }

}
     /* ============================
      Events
  ============================ */
// this for home page
public async getUpcomingEvents(): Promise<IEvent[]> {

  try {

    const items = await this.sp.web.lists
      .getByTitle("MR_SL_Events")
      .items
      .select(
  "Id",
  "Title",
  "Description",
  "StartDate",
  "EndDate",
  "IsActive",
  "AttachmentFiles"
)
.expand("AttachmentFiles")
      .filter("IsActive eq 1")
      .orderBy("StartDate", true)();

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    return items
      .filter((item: any) => {

        if (!item.EndDate) {
          return true;
        }

        const endDate = new Date(item.EndDate);
        endDate.setHours(0, 0, 0, 0);

        return endDate >= today;

      })
     .map((item:any)=>{

const attachments = item.AttachmentFiles || [];

let fileUrl = "";

if (attachments.length > 0) {

    fileUrl =
      `${window.location.origin}${attachments[0].ServerRelativeUrl}`;

}

return {

    Id: item.Id,

    Title: item.Title,

    Description: item.Description,


    StartDate: item.StartDate,

    EndDate: item.EndDate,

    IsActive: item.IsActive,

    FileUrl: fileUrl,

    Attachments: attachments.map((file:any)=>({

        FileName: file.FileName,

        ServerRelativeUrl: file.ServerRelativeUrl

    }))

};

})

  } catch (error) {

    console.error("Error fetching upcoming events", error);
    return [];

  }
}
// this for view all
public async getAllEvents(): Promise<IEvent[]> {

  try {

    const items = await this.sp.web.lists
      .getByTitle("MR_SL_Events")
      .items
      .select(
        "Id",
        "Title",
        "Description",
        "StartDate",
        "EndDate",
        "IsActive",
        "AttachmentFiles"
      )
      .expand("AttachmentFiles")
      .filter("IsActive eq 1")
      .orderBy("StartDate", true)();

    return items.map((item: any) => {

      const attachments = item.AttachmentFiles || [];

      let fileUrl = "";

      if (attachments.length > 0) {

        fileUrl =
          `${window.location.origin}${attachments[0].ServerRelativeUrl}`;

      }

      return {

        Id: item.Id,

        Title: item.Title,

        Description: item.Description,

        StartDate: item.StartDate,

        EndDate: item.EndDate,

        IsActive: item.IsActive,

        FileUrl: fileUrl,

        Attachments: attachments.map((file: any) => ({

          FileName: file.FileName,

          ServerRelativeUrl: file.ServerRelativeUrl

        }))

      };

    });

  }

  catch (error) {

    console.error("Error fetching all events", error);

    return [];

  }

}
 /* ============================
      FavouriteLinks
  ============================ */
public async getFavouriteLinks(): Promise<IFavouriteLink[]> {

    try {

        const items = await this.sp.web.lists
            .getByTitle("MR_SL_QuickLinks")
            .items
            .select(
                "Id",
                "Title",
                "RedirectURL",
                "DisplayOrder",
                "IsActive"
            )
            .filter("IsActive eq 1")
            .orderBy("DisplayOrder", true)();

        return items as IFavouriteLink[];

    }

    catch (error) {

        console.log(error);

        return [];

    }

}

public async getRecentDocuments(): Promise<ISharedDocument[]> {

    try {

        let documents: ISharedDocument[] = [];

        // Get all document libraries
        const libraries = await this.sp.web.lists
            .select(
                "Id",
                "Title",
                "BaseTemplate",
                "Hidden"
            )
            .filter("BaseTemplate eq 101 and Hidden eq false")();

        for (const library of libraries) {

            try {

                const files = await this.sp.web.lists
                    .getByTitle(library.Title)
                    .rootFolder
                    .files
                    .select(
                        "Name",
                        "ServerRelativeUrl",
                        "TimeLastModified"
                    )
                    .top(20)();

                documents.push(

                    ...files.map((file: any) => ({

                        Id: 0,

                        LibraryName: library.Title,

                        FileName: file.Name,

                        FileUrl:
                            window.location.origin +
                            file.ServerRelativeUrl,

                        Modified: file.TimeLastModified,

                        FileType:
                            file.Name.split(".").pop()?.toLowerCase() || ""

                    }))

                );

            } catch {

                // Skip inaccessible libraries

            }

        }

        return documents
            .sort(
                (a, b) =>
                    new Date(b.Modified).getTime() -
                    new Date(a.Modified).getTime()
            )
            .slice(0, 4);

    }

    catch (error) {

        console.log(error);

        return [];

    }

}
public async getDiscussionBoard(): Promise<IDiscussionBoard[]> {

    try {

        const items = await this.sp.web.lists
            .getByTitle("MR_SL_DiscussionBoard")
            .items
            .select(
                "Id",
                "Title",
                "Message",
                "Created",
                "ReplyCount",
                "DisplayOrder",
                "IsActive",
                 "Author/Id",
    "Author/Title",
    "Author/EMail"
              
        )
        .expand("Author")
        .filter("IsActive eq 1 and IsReply eq 0") 
        .orderBy("Created", false)();
        //.orderBy("DisplayOrder", true)();

        return items as IDiscussionBoard[];

    }

    catch (error) {

        console.log(error);

        return [];

    }

}
public async getReplies(
    parentId:number
):Promise<IDiscussionBoard[]>{

return await this.sp.web.lists
.getByTitle("MR_SL_DiscussionBoard")
.items
.select(
"Created",
"Id",
                "Title",
                "Message",
                "IsActive",
                "Author/Id",     
"Author/Title",
"Author/EMail"
)
.expand("Author")
.filter(
`ParentId eq ${parentId} and IsReply eq 1`
)
.orderBy("Created",true)();

}

public async addReply(
    parentId: number,
    message: string
): Promise<void> {

    // Get current user
//const currentUserId = this.context.pageContext.legacyPageContext.userId;
    // Get current reply count
    const parent = await this.sp.web.lists
        .getByTitle("MR_SL_DiscussionBoard") 
        .items
        .getById(parentId)
        .select("ReplyCount")();

    const replyCount = parent.ReplyCount || 0;

    // Add reply
    await this.sp.web.lists
        .getByTitle("MR_SL_DiscussionBoard")
        .items
        .add({

            Title: "",

            Message: message,

            ParentId: parentId,

            IsReply: true,

            IsActive: true

        });

    // Update parent discussion
  await this.sp.web.lists
    .getByTitle("MR_SL_DiscussionBoard")
    .items
    .getById(parentId)
    .update({
        ReplyCount: replyCount + 1
    });

}

public async addDiscussion(
    title: string,
    message: string,
    isQuestion: boolean
): Promise<void> {

    await this.sp.web.lists
        .getByTitle("MR_SL_DiscussionBoard")
        .items
        .add({

            Title: title,

            Message: message,

            IsQuestion: isQuestion,

            IsReply: false,

            IsActive: true,

            ReplyCount: 0,


        });

}

public async updateReply(
    replyId: number,
    message: string
): Promise<void> {

    await this.sp.web.lists
        .getByTitle("MR_SL_DiscussionBoard")
        .items
        .getById(replyId)
        .update({

            Message: message

        });

}

}
