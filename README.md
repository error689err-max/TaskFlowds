import {
  SdfAccordion,
  SdfAlertBanner,
  SdfAlertInline,
  SdfAlertToast,
  SdfAvatar,
  SdfBox,
  SdfBusyIndicator,
  SdfButton,
  SdfExpandableBox,
  SdfFocusPane,
  SdfIcon,
  SdfSelectSimple,
  SdfSpotIllustration,
  SdfTab,
  SdfTabGroup,
  SdfTag,
  SdfTextarea,
  SdfTooltip,
} from "@waypoint/react-components";

import React, { useCallback, useContext, useEffect, useRef, useState } from "react";
import rehypeRaw from "rehype-raw";
import remarkGfm from "remark-gfm";
import useMountTransition from "./useMountTransition";
import { UserContext } from "../App";
import {
  fnAddFolder,
  fnChangeChatFav,
  fnChangeChatTitle,
  fnChatUpdateFolder,
  fnContinueChatAPI,
  fnDeletFolder,
  fnDeleteChatID,
  fnDeleteSuccessFile,
  fnEditFolderTitle,
  fnGetAllChatsForChatId,
  fnGetMeetingSchedulerAgentDetails,
  fnGetMyGpts,
  fnGetPresingedURL,
  fnGetStatusForDocument,
  fnGetSummaryandTranslateDocumentJobStatus,
  fnNewChatAPI,
  fnPOSTUsersOptStatus,
  fnRecentChats,
  fnScheduleMetting,
  fnSendFeedbackResponse,
  fnWebexGroupChat,
  fnWebexMeetingSummary,
  fnRecentGPT,
  fnDitaLangTranslation,
  fnExcelAgentChat,
  fnCreateNotification,
  fnOutlookAgent,
  fnConfluenceAgent
} from "../lib/APIConnects";
import "katex/dist/katex.min.css";
import Markdown from "./MarkDown";
import MarkDownReply from "./MarkDownReply";
import newpdfLimitation from "../Images/Instructions for Responsible Use of ESI Fusion.pdf";
import ADPLogo from "../Images/ADP_logo_Always.png";
import brandassitsvg from "../Images/chat_brandAssist.png";
import moment from "moment-timezone";
import "rsuite/Tree/styles/index.css";
import { FeedbackSection } from "../Components/FeedbackSection";
import { UploadGeneral } from "../Components/UploadGeneral";
import ImageChartDisplay from "../Components/ImageChartDisplay";
import { prompts_data } from "../Store/PromptsData";
import {
  setIsError,
  setStopListening,
  setTranscript
} from "../redux_slices/voiceSlice";
import SpeechRecognition, {
  useSpeechRecognition,
} from "react-speech-recognition";
import { useDispatch, useSelector } from "react-redux";

import "../css/voice.css";
import "rsuite/Dropdown/styles/index.css";
import VoiceWaveform from "../Components/VoiceWaveform";
import { useAnimation } from "framer-motion";
import InlineVoice from "../Components/InlineVoice";
import FileUploadComponent from "../Components/FileUploadComponent";
import FileDisplayCard from "../Components/FileDisplayCard";
import { MAX_FIlES } from "../helper_function/file_validations";
import VoiceChat from "../Components/VoiceChat";
import SideBar from "./SideBar/SideBar";
import { setRecentChats } from "../redux_slices/chatSlice";
import { CustomSelect } from "../Components/CustomSelect";
import { GPTscreen } from "./GPTscreen";
import Webex_Groups_Search from "../Components/WebexGroupsSearch";
import DynamicStatusMessage from "../Components/DynamicStatusMessage";
import { web_browser_svg_resp } from "../Store/SvgHub";
import Help from "../Components/Help/Help";
import { GPTs } from "./GPTScreens/GPTS";
import CreateNotification from "./CreateNotification";
import AIGI_doc_extraction_new from "../Components/AIGI_Doc";
import BPQDocumentProcessor from "../Components/BPQDocumentProcessor";
import FIJDocumentProcessor from "../Components/FIJDocumentProcessor";
import AttachmentActionMenu from "../Components/AttachmentActionMenu";
import FilterDropdownMenu from "./FilterDropdownMenu";

const AgentAssist = ({
  blocked_chat_context,
  setBlocked_chat_context,
  setBrowerSearchId,
  browerSearchId,
  folder_id,
  setFolder_id,
  inputQueryRef,
  isHome,
  setHome,
  responseValues,
  setresponseValues,
  showWelcomeMessage,
  setshowWelcomeMessage,
  newChatLoading,
  setNewChatLoading,
  document_id,
  setdocument_id,
  documentStatusList,
  setDocumentStatusList,
  chatID,
  setchatID,
  message,
  setMessage,
  setblnLeftNavigationClicked,
  isResponseDone,
  setisResponseDone,
  selectedPrompt,
  setselectedPrompt,
  isFileProcessing,
  setIsFileProcessing,
  setFileList,
  chatInputErrorMessage,
  setChatInputErrorMessage,
  loadingFromRecent,
  setloadingFromRecent,
  displayErrorMessage,
  setdisplayErrorMessage,
  isFav,
  setisFav,
  setIconStyle,
  fileUploadComponentErrors,
  setFileUploadComponentErros,
  filesToUpload,
  setFilesToUpload,
  openGPTScreen,
  setOpenGPTScreen,
  resetInlineVoiceonClickHome,
  newOpenGPTScreen,
  setNewOpenGPTScreen,
  gptMode, 
  setGptMode,
  dispAlertMessage, setdispAlertMessage,
  openAIGIFlow, setOpenAIGIFlow,
  openBPQ, setOpenBPQ,
  openFIJ, setOpenFIJ
}) => {
  const [activeKnowledgeFilters, setActiveKnowledgeFilters] = useState({
    ekm: false,
    sharepoint: false,
    confluence: false
  });

  const [helpVis, setHelpVis] = useState(false);
  const helpFeedbackRef = useRef(null);
  const helpSupportRef = useRef(null)
  const translations = useSelector((state) => state.language.translations);
  const [isEsGroup] = useState(false);
  const [sessionGroupsModal, setSessionGroupsModal] = useState({
    visible: false,
    groups: [],
    text: "",
    actions: [],
  });
  const [openSearch, setOpenSearch] = useState(false);
  const allchatsRec = useRef([]);
  const [openSideNav, setOpenSideNav] = useState(false);

  const dispatch = useDispatch();
  const searchRef_webex_groups = useRef(null);

  const audioRef = useRef(null);
  const [voice_mode, set_voice_mode] = useState("voice_inline");
  const [listeningWaveForm, setListeningWaveForm] = useState(false);
  const endOfMessageRef1 = useRef();
  const { azureUser } = useContext(UserContext);
  const { setazureUser } = useContext(UserContext);
  const searchRef = useRef(null);
  const createNoteRef = useRef(null);
  const titleRef = useRef(null);
  const foldertitleRef = useRef(null);
  const editfoldertitleRef = useRef(null);
  const feedbackRef = useRef(null);
  const [messageReadyNotificationContext, setMessageReadyNotificationContext] =
    useState({
      is_message_ready: false,
    });
  const [windowDimensions, setWindowDimensions] = useState({
    height: window.innerHeight,
    width: window.innerWidth,
  });
  const [chatsSectionNavstate, setChatsSectionNavState] = useState({
    isNavActionClicked: false,
    chat_id: null,
    Target_folder_id: null,
    Target_folder_label: null,
    action: "",
  });
  const [allFolders, setAllFolders] = useState([]);
  const [isAddFolderClicked, SetIsAddFolderClicked] = useState(false);
  const [isupdateFolderClicked, setIsupdateFolderClicked] = useState(false);
  const [isDeleteFolderClicked, setIsDeleteFolderClicked] = useState(false);
  const [deleteFolderId, setDeletedFolderId] = useState(null);
  const pdfRef = useRef(null);
  const [chatimage] = useState(true);
  const [resChatSearch, setresChatSearch] = useState("");
  const [responseValuesLimit, setResponseValuesLimit] = useState(false);
  const [chatHistoryLoading, setChatHistoryLoading] = useState(false);
  const [changeTitleError, setChangeTitleError] = useState(false);
  const [recentChats, setrecentChats] = useState(null);
  const [filteredRecentChats, setFilteredRecentChats] = useState([]);
  const [favRecentChats, setFavRecentChats] = useState([]);
  const [filteredFavRecentChats, setFilteredFavRecentChats] = useState([]);
  const [isMounted, setIsMounted] = useState(true);
  const hasTransitionedIn = useMountTransition(isMounted, 1000);
  const [segmentControlState, setSegmentControlState] = useState(1);
  const [uploadHanledClick, setuploadHanledClick] = useState(false);
  const [isDeleteClicked, setisDeleteClicked] = useState(false);
  const [deleteChatID, setdeleteChatID] = useState("");
  const [isLoading, setisLoading] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [isChangeClicked, setisChangeClicked] = useState(false);
  const [newTitle, setnewTitle] = useState("");
  const [folderTitle, setFolderTitle] = useState("");
  const [editChatTitleID, seteditChatTitleID] = useState("");
  const [isPromptClicked, setisPromptClicked] = useState(false);
  const [activeTab, setActiveTab] = useState("Basic");
  const [isPromptOpen, setIsPromptOpen] = useState(false);
  const [promptsData, setpromptsData] = useState([]);
  const [viewStaticData, setviewStaticData] = useState({
    isClicked: false,
    clickedItem: "",
    Data: [],
  });
  const [onFilePasteFn, setOnFilePasteFn] = useState(null);

  const [dispError, setdispError] = useState({
    status: false,
    message: "",
  });

  // const [dispAlertMessage, setdispAlertMessage] = useState({
  //   status: false,
  //   code: "",
  //   message: "",
  // });

  const [feedbackCategories, setfeedbackCategories] = useState([]);
  const [optOutWarning, setoptOutWarning] = useState(false);

  const [favMessageLoad, setFavMessageLoad] = useState(false);
  const prompts_Data = prompts_data(translations);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (!chatHistoryLoading && favMessageLoad) {
      setdispAlertMessage({
        status: true,
        code: "success",
        message: translations?.Favourite,
      });
      setFavMessageLoad(false);
    }
  }, [chatHistoryLoading]);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (resetInlineVoiceonClickHome.current) {
      console.log("resetInlineVoiceonClickHome changed to true");
      resetInlineVoiceonClickHome.current = false;
      resetInlineVoiceChat();
    }
  }, [resetInlineVoiceonClickHome.current]);

  let illustrationNames = {
    "Presentation Creator": "documents-pencil",
    "Explain like 5!": "teacher",
    "Excel Assist": "book",
    "Email Writer": "email",
    "Translator": "gear-arrows",
    "Summarize Documents": "notes",
    "User Story Coach": "lightbulb",
    "Meeting Scheduler": "calendar-avatar",
    "Webex Group Chat": "connections",
    "Webex Meeting Summary Agent": "notes",
  };

  const voices = window.speechSynthesis.getVoices();

  let CapabilitiesArray = [
    "Agent Assist is capable of delivering generic information and assisting with task completion, even without access to specific data points.",
    "Agent Assist comes with pre-set prompts that can aid you in addressing specific requests.",
    "You can utilize Agent Assist to spark creativity and produce fresh content."
  ];

  const fnNewChat = () => {
    localStorage.setItem("chat_id", "");
    localStorage.setItem(
      "chatsessionid",
      Number(localStorage.getItem("chatsessionid")) + 1
    );
    localStorage.setItem("LatestToRender", "");

    setshowWelcomeMessage(false);
    setHome(false);
    setNewChatLoading(true);
    setresponseValues([]);
    resetInlineVoiceChat();
    setBrowerSearchId("");    
    setOpenAIGIFlow(false);
    setOpenBPQ(false);
    setOpenFIJ(false)
    setBlocked_chat_context({
      is_blocked: false,
      blocked_message: "",
    });
    setselectedPrompt("");
    setMessage("");
    const default_model = (azureUser.models_accessible ?? []).filter(
      (model) => model.default_model
    );
    setGptMode({ is_gpt: false, details: {} })
    setazureUser((prev) => {
      return {
        ...prev,
        selected_model: {
          model_id: default_model?.[0]?.model_id,
          model_name: default_model?.[0]?.llm_name,
          document_supported: default_model?.[0]?.document_supported,
          default_model: default_model?.[0]?.default_model,
        },
      };
    });

    if (inputQueryRef && inputQueryRef.current) {
      inputQueryRef.current.value = "";
      inputQueryRef.current.focus();
    }
    setChatInputErrorMessage({
      ...chatInputErrorMessage,
      status: false,
      message: "",
    }); // scope change
    // setdisplayErrorMessage({ status: false, message: "" });
    setFileList([]);
    // setfileName("");// can remove it
    setloadingFromRecent(false); // need to change the scope
    setdisplayErrorMessage({ status: false, message: "" }); // scope change
    setDocumentStatusList([]);
    setdocument_id([]);
    setisResponseDone(true);
    setIsFileProcessing(false);
    setisFav(false);
    setFilesToUpload([]);
    setFileUploadComponentErros({ type: "", message: "", status: false });
    setchatID();
    setTimeout(() => {
      setNewChatLoading(false);
      setshowWelcomeMessage(true);
    }, 1000);
  };


  function Model_state_rest(is_gpt=false,chat=null) {
    if(is_gpt){
      setGptMode({ is_gpt: true, details: chat })
    }
    else{
      setGptMode({ is_gpt: false, details: {} })
    }
    localStorage.setItem("chat_id", "");
    localStorage.setItem(
      "chatsessionid",
      Number(localStorage.getItem("chatsessionid")) + 1
    );
    localStorage.setItem("LatestToRender", "");
    // setChatSessionId(prev => prev+1);
    setshowWelcomeMessage(false);
    setHome(false);
    setNewChatLoading(true);
    setOpenBPQ(false);
    setOpenAIGIFlow(false);
    setresponseValues([]);
    // setisLeftnaviconselected(false);
    setselectedPrompt("");
    setMessage("");

    if (inputQueryRef && inputQueryRef.current) {
      inputQueryRef.current.value = "";
      inputQueryRef.current.focus();
    }
    setChatInputErrorMessage({
      ...chatInputErrorMessage,
      status: false,
      message: "",
    }); // scope change
    // setdisplayErrorMessage({ status: false, message: "" });
    setFileList([]);
    // setfileName("");// can remove it
    setloadingFromRecent(false); // need to change the scope
    setdisplayErrorMessage({ status: false, message: "" }); // scope change
    setDocumentStatusList([]);
    setBlocked_chat_context({
      is_blocked: false,
      blocked_message: "",
    });
    setdocument_id([]);
    setisResponseDone(true);
    setIsFileProcessing(false);
    setisFav(false);
    setFilesToUpload([]);
    setFileUploadComponentErros({ type: "", message: "", status: false });
    setchatID();
    setTimeout(() => {
      setNewChatLoading(false);
      setshowWelcomeMessage(true);
    }, 1000);
  }
  console.log(azureUser?.esid);

  // Global key listener
  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    setTimeout(() => inputQueryRef?.current?.focus(), 500);
  }, []);

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    const GlobalkeyHandler = (e) => {
      // Allow default copy/paste behavior
      if (
        (e.ctrlKey || e.metaKey) &&
        (e.key === "c" || e.key === "v" || e.key === "x")
      ) {
        return;
      }

      const active = document.activeElement;
      if (
        (createNoteRef.current && (active === createNoteRef.current || createNoteRef.current.contains(active))) ||
        (searchRef.current && active === searchRef.current) ||
        (searchRef_webex_groups.current &&
          active === searchRef_webex_groups.current) ||
        (titleRef.current && active === titleRef.current) ||
        (foldertitleRef.current && active === foldertitleRef.current) ||
        (editfoldertitleRef.current && active === editfoldertitleRef.current) ||
        (feedbackRef.current && active === feedbackRef.current) ||
        (helpFeedbackRef.current && active === helpFeedbackRef.current) ||
        (helpSupportRef.current && active === helpSupportRef.current)
      ) {
        return;
      }

      if (
        !e.ctrlKey &&
        !e.metaKey &&
        !e.altKey &&
        !e.shiftKey &&
        e.key.length === 1 &&
        !helpVis
      ) {
        inputQueryRef.current?.focus();
      }
    };

    document.addEventListener("keydown", GlobalkeyHandler);
    return () => {
      document.removeEventListener("keydown", GlobalkeyHandler);
    };
  }, [helpVis]);

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    window.addEventListener("resize", getWindowDimensions);
    getWindowDimensions();
    setpromptsData(prompts_Data);
  }, []);

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    fnHandleFilterRecentchats();
  }, [resChatSearch, segmentControlState]);

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    setpromptsData(prompts_Data);
  }, [translations])

  const registerPasteHandler = useCallback((fn) => {
    setOnFilePasteFn(() => fn);
  }, []);

  const [baseFontSize, setBaseFontSize] = useState(() => {
    const savedSize = localStorage.getItem('appFontSize');
    return savedSize ? parseInt(savedSize, 10) : 16;
  });

  const decreaseFontSize = () => {
    const currentSize = parseInt(localStorage.getItem('appFontSize') || '16', 10);
    const newFontSize = Math.max(currentSize - 1, 14);
    if (newFontSize !== currentSize) {
      setBaseFontSize(newFontSize);
      document.documentElement.style.setProperty('--global-font-scale', `${newFontSize / 16}`);
      localStorage.setItem('appFontSize', newFontSize.toString());
      console.log("Font size decreased to:", newFontSize);
    } else {
      console.log("Font size already at minimum (14)");
    }
  };

  const resetFontSize = () => {
    const defaultSize = 16;
    setBaseFontSize(defaultSize);
    document.documentElement.style.setProperty('--global-font-scale', '1');
    localStorage.setItem('appFontSize', defaultSize.toString());
  };

  const IncreaseFontSize = () => {
    const currentSize = parseInt(localStorage.getItem('appFontSize') || '16', 10);
    const newFontSize = Math.min(18, currentSize + 1);
    if (newFontSize !== currentSize) {
      setBaseFontSize(newFontSize);
      document.documentElement.style.setProperty('--global-font-scale', `${newFontSize / 16}`);
      localStorage.setItem('appFontSize', newFontSize.toString());
      console.log("Font size increased to:", newFontSize);
    } else {
      console.log("Font size already at maximum (18)");
    }
  };

  const modeF = localStorage.getItem('dark-mode') ?? "one";
  const [darkModeVars, setDarkModeVars] = useState({
    homeBoxes: modeF === 'one' ? "shadowed" : "filled",
    sideBarColor: modeF === 'one' ? "" : "#252943",
    darkBtnIcon: modeF === 'one' ? "action-lighten" : "action-darken",
    backgroundChatUser: modeF === 'one' ? "" : "#121C4E",
    webSearchBg: modeF === "one" ? "#d1d5db" : "#171A28",
  })

  const fnChangeMode = () => {
    const buttonContainer = document?.querySelector("sfc-shell-app-bar")
      ?.shadowRoot?.querySelector(".logged-in-user")
      ?.querySelector("sdf-floating-pane")
      ?.querySelector(".avatar-menu");
    const darkModeBtn = buttonContainer.querySelector("button[data-btn-action='darkMode']");
    const tabArrow = document.querySelector(".tab-dark-mode-grp")?.shadowRoot?.querySelector("sdf-scroller")?.shadowRoot?.querySelectorAll(".next, .prev")


    if (window.SynergConfig.mode === 'one') {
      localStorage.setItem("dark-mode", "one-dark")
      window.SynergConfig.setMode('one-dark');
      setDarkModeVars(prev => ({ ...prev, homeBoxes: "filled", sideBarColor: "#252943", darkBtnIcon: "action-darken", backgroundChatUser: "#121C4E", webSearchBg: "#171A28" }));
      if (darkModeBtn) {
        darkModeBtn.innerHTML = '<sdf-icon icon="action-lighten" style="font-size: inherit !important;margin-right: 0px !important;"></sdf-icon>'
        document.documentElement.style.setProperty('--lg-txt-color', '#fff');
        document.documentElement.style.setProperty('--main-app-background', '#171a28');
        document.documentElement.style.setProperty('--home-box-button', '#fff');
        document.documentElement.style.setProperty('--home-global-dropdown', '#252943');
        document.documentElement.style.setProperty('--res-table-bg-th', '#393533');
        document.documentElement.style.setProperty('--res-table-border-b', '#938c85');
        document.documentElement.style.setProperty('--shimmer-text-color', '#8d8585');
        document.documentElement.style.setProperty('--shimmer-bg-color', 'rgba(255,255,255,0.1)');
        document.documentElement.style.setProperty('--small-icon-color', '#8a96a7');
        document.documentElement.style.setProperty('--textarea-icon-hover-color', '#171A28');
        document.documentElement.style.setProperty('--small-icon-bg', '#303659');
        document.documentElement.style.setProperty('--faq-txtarea-bg', '#282D47');
        document.documentElement.style.setProperty('--chat-faq-background', '#3B4162');
        document.documentElement.style.setProperty('--chat-bubble-background', '#252943');
        document.documentElement.style.setProperty('--drop-down-background', '#252943');
      }
      if (tabArrow && tabArrow.length > 0) {
        tabArrow.forEach((ele) => {
          ele.style.background = "#171a28"
        })
      }
    } else {
      localStorage.setItem("dark-mode", "one")
      window.SynergConfig.setMode('one');
      document.documentElement.style.setProperty('--lg-txt-color', 'black');
      document.documentElement.style.setProperty('--faq-txtarea-bg', '#fff');
      document.documentElement.style.setProperty('--shimmer-bg-color', 'rgba(224,224,224,0.7)');
      document.documentElement.style.setProperty('--shimmer-text-color', '#222');
      document.documentElement.style.setProperty('--main-app-background', 'white');
      document.documentElement.style.setProperty('--chat-bubble-background', '#f2f0ed');
      document.documentElement.style.setProperty('--chat-faq-background', '#f2f0ed');
      document.documentElement.style.setProperty('--home-box-button', '#3252ad');
      document.documentElement.style.setProperty('--home-global-dropdown', 'white');
      document.documentElement.style.setProperty('--res-table-bg-th', '#00000019');
      document.documentElement.style.setProperty('--res-table-border-b', '#00000019');
      document.documentElement.style.setProperty('--small-icon-color', '#475569');
      document.documentElement.style.setProperty('--textarea-icon-hover-color', 'gainsboro');
      document.documentElement.style.setProperty('--small-icon-bg', '#d1d5db');
      document.documentElement.style.setProperty('--drop-down-background', 'white');
      setDarkModeVars(prev => ({ ...prev, homeBoxes: "shadowed", sideBarColor: "", darkBtnIcon: "action-lighten", backgroundChatUser: "", webSearchBg: "#d1d5db" }));
      if (darkModeBtn) {
        darkModeBtn.innerHTML = '<sdf-icon icon="action-darken" style="font-size: inherit !important;margin-right: 0px !important;"></sdf-icon>'
      }
      if (tabArrow && tabArrow.length > 0) {
        tabArrow.forEach((ele) => {
          ele.style.background = ""
        })
      }
    }
  }

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {

    if (newOpenGPTScreen) {
      setTimeout(() => {
        const tabArrow = document.querySelector(".tab-dark-mode-grp")?.shadowRoot?.querySelector("sdf-scroller")?.shadowRoot?.querySelectorAll(".next, .prev")
        if (tabArrow && tabArrow.length > 0) {
          if (localStorage.getItem('dark-mode') === 'one-dark') {
            tabArrow.forEach((ele) => {
              ele.style.background = "#171a28"
            })
          }
        }
      }, 200)
    }

  }, [newOpenGPTScreen])

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    document.documentElement.style.setProperty('--global-font-scale', `${baseFontSize / 16}`);

    //dark mode
    if (localStorage.getItem('dark-mode') === 'one-dark') {
      document.documentElement.style.setProperty('--lg-txt-color', '#fff');
      document.documentElement.style.setProperty('--faq-txtarea-bg', '#282D47');
      document.documentElement.style.setProperty('--shimmer-bg-color', 'rgba(255,255,255,0.1)');
      document.documentElement.style.setProperty('--shimmer-text-color', '#8d8585');
      document.documentElement.style.setProperty('--chat-bubble-background', '#252943');
      document.documentElement.style.setProperty('--chat-faq-background', '#3B4162');
      document.documentElement.style.setProperty('--main-app-background', '#171a28');
      document.documentElement.style.setProperty('--home-box-button', '#fff');
      document.documentElement.style.setProperty('--home-global-dropdown', '#252943');
      document.documentElement.style.setProperty('--res-table-bg-th', '#393533');
      document.documentElement.style.setProperty('--res-table-border-b', '#938c85');
      document.documentElement.style.setProperty('--small-icon-color', '#8a96a7');
      document.documentElement.style.setProperty('--textarea-icon-hover-color', '#171A28');
      document.documentElement.style.setProperty('--small-icon-bg', '#303659');
      document.documentElement.style.setProperty('--drop-down-background', '#252943');
    }


  }, []);

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    console.log("basefont size var changed : ", baseFontSize)
    try {
      const buttonContainer = document.querySelector("sfc-shell-app-bar")
        ?.shadowRoot?.querySelector(".logged-in-user")
        ?.querySelector("sdf-floating-pane")
        ?.querySelector(".avatar-menu");
      if (buttonContainer) {
        // Update decrease button (A-)
        const decreaseBtn = buttonContainer.querySelector("button[data-font-action='decrease']");
        if (decreaseBtn) {
          decreaseBtn.disabled = baseFontSize <= 14;
          decreaseBtn.style.backgroundColor = baseFontSize <= 14 ? '#9ca3af' : '#476bc3';
          decreaseBtn.style.cursor = baseFontSize <= 14 ? 'not-allowed' : 'pointer';
        }

        // Update reset button (A)
        const resetBtn = buttonContainer.querySelector("button[data-font-action='reset']");
        if (resetBtn) {
          resetBtn.disabled = baseFontSize === 16;
          resetBtn.style.backgroundColor = baseFontSize === 16 ? '#9ca3af' : '#476bc3';
          resetBtn.style.cursor = baseFontSize === 16 ? 'not-allowed' : 'pointer';
        }

        // Update increase button (A+)
        const increaseBtn = buttonContainer.querySelector("button[data-font-action='increase']");
        if (increaseBtn) {
          increaseBtn.disabled = baseFontSize >= 18;
          increaseBtn.style.backgroundColor = baseFontSize >= 18 ? '#9ca3af' : '#476bc3';
          increaseBtn.style.cursor = baseFontSize >= 18 ? 'not-allowed' : 'pointer';
        }
      }
    } catch (error) {
      console.log("Font buttons not yet available");
    }
  }, [baseFontSize]);

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    const buttonContainer = document.querySelector("sfc-shell-app-bar")
      ?.shadowRoot?.querySelector(".logged-in-user")
      ?.querySelector("sdf-floating-pane")
      ?.querySelector(".avatar-menu");

    const isFontBtn = buttonContainer?.querySelectorAll("button[data-font-action]");

    if (isFontBtn) {
      if (window.devicePixelRatio > 1.2) {
        isFontBtn.forEach(btn => btn.style.fontSize = "9px");
      } else {
        isFontBtn?.forEach(btn => btn.style.fontSize = "11px");
      }
    }
  }, [window.devicePixelRatio]);

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    setTimeout(() => {
      let useravatarContainer = document
        .querySelector("sfc-shell-app-bar")
        .shadowRoot.querySelector(".logged-in-user")
        .querySelector("sdf-floating-pane")
        .querySelector(".avatar-menu");

      let isToggleSwitchPresent =
        useravatarContainer.querySelector("#optToggle");

      let useravatar =
        useravatarContainer.lastElementChild.querySelector("ul");

      const userIcon = useravatarContainer?.querySelector("sdf-avatar")
      const userName = useravatarContainer?.querySelector("h2")
      const userRole = useravatarContainer?.querySelector("h3")
      const toogleFont = isToggleSwitchPresent?.shadowRoot?.querySelector("span")?.querySelector("label");
      const SdfButtonFont = useravatar.querySelector("li")?.querySelector("sdf-button")
      if (userIcon) {
        userIcon.style.fontSize = `calc(var(--global-font-scale) * 1.5rem)`;
        userIcon.style.height = `calc(var(--global-font-scale) * 4rem)`;
        userIcon.style.width = `calc(var(--global-font-scale) * 4rem)`;
      }
      if (userName) {
        userName.style.fontSize = `calc(var(--global-font-scale) * 1.5rem)`;
      }
      if (userRole) {
        userRole.style.fontSize = `calc(var(--global-font-scale) * 1.125rem)`;
      }
      if (toogleFont) {
        toogleFont.style.fontSize = `calc(var(--global-font-scale) * 1rem)`;
      }
      if (SdfButtonFont) {
        SdfButtonFont.style.fontSize = `calc(var(--global-font-scale) * 1rem)`;
      }

      if (
        isToggleSwitchPresent === null ||
        isToggleSwitchPresent === undefined
      ) {
        console.log("Not present");


        let li = document.createElement("li");
        console.log(useravatar);
        let toggle = document.createElement("sdf-switch");
        toggle.setAttribute("id", "optToggle");

        if (azureUser.isPreviledgedUser) {
          toggle.setAttribute("disabled", "true");
          toggle.setAttribute("label", "Opt-out");
        } else {
          toggle.setAttribute("label", "Opt-in");
          toggle.setAttribute("checked", "true");
        }

        // toggle.style.padding = "5px";
        toggle.style.minWidth = "15rem";
        toggle.style.display = "flex";
        toggle.style.justifyContent = "center";
        li.appendChild(toggle);
        useravatar.insertBefore(li, useravatar.firstChild);

        toggle.addEventListener("sdfChange", (e) => {
          fnChangedOptOptions(e);
          e.preventDefault();
          console.log(e);
        });
      }
      //   Copied commented by Umesh for future release
      let isFontDivPresent = useravatarContainer.querySelector("#font-button-div");
      if (isFontDivPresent === null ||
        isFontDivPresent === undefined) {
        let li1 = document.createElement("li");
        let box = document.createElement("div");
        box.setAttribute("id", "font-button-div")
        box.style.cssText = `        
          display:flex;
            gap:0.5rem;
            justify-content: center;
            margin-bottom: 1rem;
          `;

        const buttonCss = `
            border: none;
            color: white;
            padding: 8px;
            border-radius: 8px;
            font-weight: 600;
            font-size: ${window.devicePixelRatio > 1.2 ? '9px' : '11px'};
            transition: all 0.3s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            min-width: 40px;
            margin: 0
          `;

        let button1 = document.createElement("button");
        button1.innerHTML = "A-";
        button1.style.cssText = buttonCss;
        button1.setAttribute('data-font-action', 'decrease');
        button1.addEventListener('click', decreaseFontSize);
        const currentSize = parseInt(localStorage.getItem('appFontSize') || '16', 10);
        button1.disabled = currentSize <= 14;
        button1.style.backgroundColor = currentSize <= 14 ? '#9ca3af' : '#476bc3';
        button1.style.cursor = currentSize <= 14 ? 'not-allowed' : 'pointer';

        let button2 = document.createElement("button");
        button2.innerHTML = "A";
        button2.style.cssText = buttonCss;
        button2.setAttribute('data-font-action', 'reset');
        button2.addEventListener('click', resetFontSize);
        button2.disabled = currentSize === 16;
        button2.style.backgroundColor = currentSize === 16 ? '#9ca3af' : '#476bc3';
        button2.style.cursor = currentSize === 16 ? 'not-allowed' : 'pointer';

        let button3 = document.createElement("button");
        button3.innerHTML = "A+";
        button3.style.cssText = buttonCss;
        button3.setAttribute('data-font-action', 'increase');
        button3.addEventListener('click', IncreaseFontSize);
        button3.disabled = currentSize >= 18;
        button3.style.backgroundColor = currentSize >= 18 ? '#9ca3af' : '#476bc3';
        button3.style.cursor = currentSize >= 18 ? 'not-allowed' : 'pointer';

        let darkModeBtn = document.createElement("button");
        darkModeBtn.innerHTML = `<sdf-icon icon="${localStorage.getItem('dark-mode') === 'one' ? "action-darken" : "action-lighten"}" style="font-size: inherit !important;margin-right: 0px !important;"></sdf-icon>`;
        darkModeBtn.style.cssText = buttonCss;
        darkModeBtn.setAttribute('data-btn-action', 'darkMode');
        darkModeBtn.addEventListener('click', fnChangeMode);
        darkModeBtn.style.backgroundColor = '#476bc3';
        darkModeBtn.style.cursor = 'pointer';

        box.appendChild(button1)
        box.appendChild(button2)
        box.appendChild(button3)
        box.appendChild(darkModeBtn)

        li1.appendChild(box);
        useravatar.insertBefore(li1, useravatar.firstChild);
      }

      let adpAppBarShadowRoot =
        document.querySelector("sfc-shell-app-bar").shadowRoot;
      let logo = adpAppBarShadowRoot.querySelector("sdf-icon");
      // logo.style.margin = "0px 0px 0px 60px";
      setIconStyle(true);
      console.log("shadowroot", adpAppBarShadowRoot);
      console.log(logo);
    }, 2000);

    setTimeout(() => {
      let adpAppBarShadowRoot =
        document.querySelector("sfc-shell-app-bar").shadowRoot;
      let logo = adpAppBarShadowRoot.querySelector("sdf-icon");
      // logo.style.margin = "0px 0px 0px 60px";
      setIconStyle(true);
      console.log("shadowroot", adpAppBarShadowRoot);
      console.log(logo);
    }, 1000);

    setTimeout(() => {
      let adpAppBarShadowRoot =
        document.querySelector("sfc-shell-app-bar").shadowRoot;
      let logo = adpAppBarShadowRoot.querySelector("sdf-icon");
      // logo.style.margin = "0px 0px 0px 60px";
      setIconStyle(true);
      console.log("shadowroot", adpAppBarShadowRoot);
      console.log(logo);
    }, 2000);

    console.log("TimeZone", Intl.DateTimeFormat().resolvedOptions().timeZone);

    fnLoadRecentChats(azureUser.accessToken);
    // fnHanldeUserDocuments(azureUser.accessToken);
  }, []);

  function fnHandleTimeZone(timestamp) {
    let t = timestamp?.toString();
    let timezone = Intl.DateTimeFormat().resolvedOptions().timeZone?.toString();
    const start = moment.tz(t, "UTC"); // original timezone
    return start.tz(timezone).format("YYYY-MM-DD hh:mm A");
  }

  function fnChangedOptOptions(e) {
    console.log("fn change opt", e);
    console.log(e.target.value);
    console.log(e.target.checked);
    console.log(e.detail);

    if (!e.target.checked) {
      setoptOutWarning(true);
    }
  }

  // eslint-disable-next-line react-hooks/exhaustive-deps
  useEffect(() => {
    if (displayErrorMessage.status)
      setdisplayErrorMessage({
        ...displayErrorMessage,
        status: false,
        message: "",
      });

    if (chatInputErrorMessage.status) {
      setChatInputErrorMessage({
        ...chatInputErrorMessage,
        status: false,
        message: "",
      });
    }
  }, [isHome]);

  function applyDynamicFileStyles() {
    const styles =
      window.devicePixelRatio > 1.13
        ? {
          "--file-display-bg": "#f0f0f066",
          "--file-display-border-radius": "6px",
          "--file-display-border-width": "1px",
          "--file-display-padding": "8px",
          "--file-display-margin-bottom": "3px",
          "--file-icon-margin-right": "2px",
          "--file-icon-font-size": "18px",
          "--file-name-font-weight": "500",
          "--file-name-font-size": "8px",
          "--file-name-margin-right": "10px",
          "--file-size-margin-left": "25px",
          "--file-size-font-size": "10px",
          "--file-size-color": "#666",
          "--file-spinner-width": "18px",
          "--file-spinner-border": "2px",
          "--file-display-margin-right": "43px",
          "--progress-ring-size": "28px",
          "--model-tag-line-size": "8px",
          "--browser-icon-width": "20px",
          "--browser-icon-height": "15px",
        }
        : {
          "--file-display-bg": "#f0f0f066",
          "--file-display-border-radius": "6px",
          "--file-display-border-width": "1px",
          "--file-display-padding": "10px",
          "--file-display-margin-bottom": "5px",
          "--file-icon-margin-right": "4px",
          "--file-icon-font-size": "22px",
          "--file-name-font-weight": "500",
          "--file-name-font-size": "11px",
          "--file-name-margin-right": "15px",
          "--file-size-margin-left": "30px",
          "--file-size-font-size": "12px",
          "--file-size-color": "#666",
          "--file-spinner-width": "26px",
          "--file-spinner-border": "3px",
          "--file-display-margin-right": "48px",
          "--progress-ring-size": "24px",
          "--model-tag-line-size": "10px",
          "--browser-icon-width": "25px",
          "--browser-icon-height": "20px",
        };

    Object.entries(styles).forEach(([key, value]) => {
      document.documentElement.style.setProperty(key, value);
    });
  }

  function getWindowDimensions() {
    console.log(window);
    console.log(window.devicePixelRatio, window.devicePixelRatio >= 1.5);
    console.log("98." + Math.random().toFixed(2) * 100 + "%");
    //for enabling scroll bar
    setTimeout(() => {
      document.getElementsByClassName("chatDocuments")[0].style.width =
        "98." + Math.random().toFixed(2) * 100 + "%";
      document.body.style.overflow = "hidden";
      document.body.style.overflow = "auto";
    }, 1000);
    if (window.devicePixelRatio > 1.13) {
      if (window.innerHeight > 650) {
        document.documentElement.style.setProperty("--help-support-height", "64vh");
        document.documentElement.style.setProperty("--help-modal-height", "70vh");
      }
      else {
        document.documentElement.style.setProperty("--help-support-height", "72vh");
        document.documentElement.style.setProperty("--help-modal-height", "79vh");
      }
      document.documentElement.style.setProperty("--nav-icons", "20px");
      document.documentElement.style.setProperty("--adp-logo", "30px");
      document.documentElement.style.setProperty("--box-size", "11px");
      document.documentElement.style.setProperty("--disclaimer-size", "9px");
      document.documentElement.style.setProperty(
        "--prompt-border-size",
        "1.5px"
      );
      document.documentElement.style.setProperty("--pdf-icon-size", "14px");
      document.documentElement.style.setProperty("--chat-display-size", "12px");
      document.documentElement.style.setProperty("--timestamp-size", "9px");
      document.documentElement.style.setProperty("--voice-icons", "22px");
      // Dropdown CSS variables for zoom levels (150% zoom) - Smaller sizing
      document.documentElement.style.setProperty(
        "--dropdown-button-width",
        "1.8rem"
      );
      document.documentElement.style.setProperty(
        "--dropdown-button-height",
        "1.4rem"
      );
      document.documentElement.style.setProperty(
        "--dropdown-button-padding",
        "0.3rem"
      );
      document.documentElement.style.setProperty(
        "--dropdown-button-margin-top",
        "-0.6rem"
      );
      document.documentElement.style.setProperty(
        "--dropdown-border-radius",
        "3px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-menu-border-radius",
        "5px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-menu-min-width",
        "140px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-menu-padding",
        "4px"
      );
      document.documentElement.style.setProperty("--dropdown-item-gap", "8px");
      document.documentElement.style.setProperty(
        "--dropdown-item-padding-vertical",
        "8px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-item-padding-horizontal",
        "10px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-font-size",
        "11px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-icon-size",
        "12px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-menu-min-width-mobile",
        "120px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-font-size-mobile",
        "10px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-item-padding-vertical-mobile",
        "6px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-item-padding-horizontal-mobile",
        "8px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-item-gap-mobile",
        "6px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-icon-size-mobile",
        "10px"
      );
      document.documentElement.style.setProperty("--font-icon-size", "14px");
    } else {
      if (window.innerHeight > 1000) {
        document.documentElement.style.setProperty("--help-support-height", "62vh");
        document.documentElement.style.setProperty("--help-modal-height", "68vh");
      }
      else {
        document.documentElement.style.setProperty("--help-support-height", "68vh");
        document.documentElement.style.setProperty("--help-modal-height", "76vh");
      }
      document.documentElement.style.setProperty("--nav-icons", "24px");
      document.documentElement.style.setProperty("--adp-logo", "40px");
      document.documentElement.style.setProperty("--box-size", "16px");
      document.documentElement.style.setProperty("--disclaimer-size", "12px");
      document.documentElement.style.setProperty("--prompt-border-size", "2px");
      document.documentElement.style.setProperty("--pdf-icon-size", "18px");
      document.documentElement.style.setProperty("--chat-display-size", "16px");
      document.documentElement.style.setProperty("--timestamp-size", "11px");
      document.documentElement.style.setProperty("--voice-icons", "32px");

      // Dropdown CSS variables for normal zoom (100% zoom)
      document.documentElement.style.setProperty(
        "--dropdown-button-width",
        "2.5rem"
      );
      document.documentElement.style.setProperty(
        "--dropdown-button-height",
        "2rem"
      );
      document.documentElement.style.setProperty(
        "--dropdown-button-padding",
        "0.5rem"
      );
      document.documentElement.style.setProperty(
        "--dropdown-button-margin-top",
        "-1rem"
      );
      document.documentElement.style.setProperty(
        "--dropdown-border-radius",
        "4px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-menu-border-radius",
        "8px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-menu-min-width",
        "200px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-menu-padding",
        "8px"
      );
      document.documentElement.style.setProperty("--dropdown-item-gap", "12px");
      document.documentElement.style.setProperty(
        "--dropdown-item-padding-vertical",
        "12px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-item-padding-horizontal",
        "16px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-font-size",
        "14px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-icon-size",
        "16px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-menu-min-width-mobile",
        "180px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-font-size-mobile",
        "13px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-item-padding-vertical-mobile",
        "10px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-item-padding-horizontal-mobile",
        "14px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-item-gap-mobile",
        "10px"
      );
      document.documentElement.style.setProperty(
        "--dropdown-icon-size-mobile",
        "14px"
      );
      document.documentElement.style.setProperty("--font-icon-size", "18px");
    }

    applyDynamicFileStyles();

    const { innerWidth: width, innerHeight: height } = window;
    console.log(width, height);

    if (width < 768) {
      setTimeout(() => {
        let appbar = document.querySelector("sfc-shell-app-bar").shadowRoot;
        let brand = appbar.querySelector(".branding-area");
        brand.style.display = "inline-flex";
      }, [1000]);
    }
    setWindowDimensions({
      width,
      height,
    });
  }

  function getJobType(taskLabel) {
    let tasklabel = taskLabel ?? "";

    switch (tasklabel) {
      case "SUMMARY":
        tasklabel = "Summarization";
        break;
      case "TRANSLATION":
        tasklabel = "Translation";
        break;
    }

    return tasklabel;
  }

  function new_Chat_response_renderer_streaming(
    streaming_response,
    requestsession
  ) {
    if (localStorage.getItem("chat_id")) {
      if (localStorage.getItem("chat_id") !== streaming_response.chat_id) {
        console.log(
          "return form new api, 'newchat->existing",
          localStorage.getItem("chat_id")
        );
        console.log("entered flash message context");
        if (streaming_response.status === "Completed") {
          if (streaming_response)
            setMessageReadyNotificationContext({
              is_message_ready: true,
              ...streaming_response,
            });
        }
        fnLoadRecentChats(azureUser.accessToken);
        return;
      }
    } else {
      setfeedbackCategories(
        streaming_response.metadata?.feedback_categories ?? []
      );
      setTimeout(() => {
        console.log(responseValues);
        //!localStorage.getItem("chat_id")
        if (
          localStorage.getItem("chat_id") === streaming_response.chat_id ||
          requestsession === localStorage.getItem("chatsessionid")
        ) {
          let chat_response_msg = {
            content:
              streaming_response.status === "Processing"
                ? emojiWrap(streaming_response.stream_message.stream_message_last_title)
                : streaming_response.message.llmResponse.content,
            isEmail: false,
            isMeeting: false,
            sender: "assistant",
            upVote: null,
            feedback: "",
            id: streaming_response.message.llmResponse.message_id,
            job_details: {
              job_id: streaming_response?.message?.job_id ?? null,
              status: "in_progress",
              phaseTracker: streaming_response?.message?.job_id
                ? getJobType(streaming_response?.message?.task_label)
                : "",
            },
            stream_check:
              streaming_response.status === "Processing" ? true : false,
            chat_id: streaming_response?.chat_id,
            timestamp: streaming_response.message.llmResponse.timestamp,
            newTimeStamp: fnHandleTimeZone(
              streaming_response.message.llmResponse.timestamp
            ),
          };
          console.log("entered new _Rendering", chat_response_msg);
          setresponseValues((prev) => {
            prev.splice(-1, 1);
            return [...prev, chat_response_msg];
          });
          setTimeout(() => {
            if (endOfMessageRef1 && endOfMessageRef1.current) {
              endOfMessageRef1.current.scrollIntoView({ behavior: "smooth" });
            }
          }, 250);

          if (
            streaming_response.status === "Completed" &&
            streaming_response.session_response &&
            Array.isArray(streaming_response.session_response.data) &&
            streaming_response.session_response.data.length > 0
          ) {
            chat_response_msg.webex_groups =
              streaming_response.session_response;
            setSessionGroupsModal({
              visible: true,
              groups: streaming_response.session_response.data,
              text: streaming_response.session_response.text,
              actions: streaming_response.session_response.actions,
            });
          }

          if (streaming_response.status === "Completed") {
            setchatID(streaming_response.chat_id);
            localStorage.setItem("chat_id", streaming_response.chat_id);

            fnLoadRecentChats(azureUser.accessToken);
            setisResponseDone(true);

            // resolve(streaming_response);
          }
        } else {
          if (streaming_response.status === "Completed") {
            console.log(
              "entered flash message context for new streaming chat with out id"
            );
            if (streaming_response)
              setMessageReadyNotificationContext({
                is_message_ready: true,
                ...streaming_response,
              });
          }
        }
      }, 1000);
    }
  }

  function emojiWrap(text) {
    const emojiRegex = /[\p{Emoji}\u200d]+/gu;
    const emoji = text.match(emojiRegex)?.[0];
    const rest = text.replace(emoji, "");
    return `<span class="emoji">${emoji}</span> ${rest}`;
  }

  function existing_chat_response_renderer_streaming(streaming_response) {
    if (localStorage.getItem("chat_id") === streaming_response?.chat_id) {
      if (
        streaming_response.message.message_id === 1 &&
        streaming_response.status === "Completed"
      ) {
        //setchatID(streaming_response.chat_id);
        fnLoadRecentChats(azureUser.accessToken);
      }

      let streaming_chat_msg = {
        content:
          streaming_response.status === "Processing"
            ? emojiWrap(streaming_response.stream_message.stream_message_last_title)
            : streaming_response.message.llmResponse.content,
        isEmail: false,
        isMeeting: false,
        sender: "assistant",
        upVote: null,
        feedback: "",
        id: streaming_response.message.llmResponse.message_id,
        job_details: {
          job_id: streaming_response?.message?.job_id ?? null,
          status: "in_progress",
          phaseTracker: streaming_response?.message?.job_id
            ? getJobType(streaming_response?.message?.task_label)
            : "",
        },
        stream_check: streaming_response.status === "Processing" ? true : false,

        timestamp: streaming_response.message.llmResponse.timestamp,
        newTimeStamp: fnHandleTimeZone(
          streaming_response.message.llmResponse.timestamp
        ),
      };

      setresponseValues((prev) => {
        prev.splice(-1, 1);
        return [...prev, streaming_chat_msg];
      });
      setTimeout(() => {
        if (endOfMessageRef1 && endOfMessageRef1.current) {
          endOfMessageRef1.current.scrollIntoView({ behavior: "smooth" });
        }
      }, 250);

      if (
        streaming_response.status === "Completed" &&
        streaming_response.session_response &&
        Array.isArray(streaming_response.session_response.data) &&
        Array.isArray(streaming_response.session_response.actions) &&
        (streaming_response.session_response.data.length > 0 ||
          streaming_response.session_response.actions.length > 0)
      ) {
        streaming_chat_msg.webex_groups = streaming_response.session_response;
        if (streaming_response.session_response.data.length > 0)
          setSessionGroupsModal({
            visible: true,
            groups: streaming_response.session_response.data,
            text: streaming_response.session_response.text,
            actions: streaming_response.session_response.actions,
          });
      }
      if (streaming_response.status === "Completed") {
        // setchatID(streaming_response.chat_id);
        localStorage.setItem("chat_id", streaming_response.chat_id);
        if (responseValues.length === 3)
          fnLoadRecentChats(azureUser.accessToken);
        setisResponseDone(true);

        // resolve(streaming_response);
      }
    } else {
      if (streaming_response) {
        if (streaming_response.status === "Completed") {
          setMessageReadyNotificationContext({
            is_message_ready: true,
            ...streaming_response,
          });
        }
      }
    }

    if (responseValues.length === 76) {
      console.log("entered to response limit");
      setResponseValuesLimit(true);
      setdisplayErrorMessage({
        ...displayErrorMessage,
        status: true,
        message: translations.message_Limit
      });
    }
  }

  function stream_response_handler(
    response,
    requestsession,
    chat_session = "New_Chat"
  ) {
    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = "";
    function read() {
      return reader.read().then(({ done, value }) => {
        let streaming_response = {};
        if (done) {
          console.log(done, "status done");
          return streaming_response;
        }

        buffer += decoder.decode(value, { stream: true });
        console.log(buffer);
        const lines = buffer.split("\n");
        buffer = lines.pop(); // Save incomplete line for next chunk
        for (const line of lines) {
          //  console.log("lines", line);
          if (line.startsWith("data:")) {
            try {
              streaming_response = JSON.parse(line.replace("data: ", ""));
              console.log(streaming_response.session_response);
              if (chat_session === "New_Chat")
                new_Chat_response_renderer_streaming(
                  streaming_response,
                  requestsession
                );
              if (chat_session === "existing_Chat") {
                existing_chat_response_renderer_streaming(streaming_response);
              }
            } catch (e) {
              console.log(e);
            }
          }
        }
        return read();
      });
    }
    return read();
  }

  function fnHandleSubmit(
    requestsession,
    selectedPrompt = "",
    searchCriteria,
    queryPrompt
  ) {
    console.log(
      "requestsession",
      requestsession,
      selectedPrompt,
      inputQueryRef
    );
    localStorage.setItem("LatestToRender", "");
    setblnLeftNavigationClicked(false);
    let queryParams = "";
    queryParams = inputQueryRef?.current?.value;
    switch (searchCriteria) {
      case "Agent Starters":
        queryParams = queryPrompt
        break;
      case "Summarize Documents":
        queryParams = queryPrompt;
        setisPromptClicked(false);
        break;
      case "Webex_Groups_Session":
        queryParams = queryPrompt;
        setSessionGroupsModal({
          visible: false,
          groups: [],
          text: "",
          actions: [],
        });
        break;
      default:
        queryParams = inputQueryRef?.current.value?.includes(
          "<target language>"
        )
          ? inputQueryRef.current.value.replaceAll("<target language>", "")
          : inputQueryRef.current.value;
    }

    // console.log("queryParams", inputQueryRef.current.value);
    if (queryParams.trim() === "") {
      setisResponseDone(true);
      setdisplayErrorMessage({
        status: true,
        message: translations.valid_Query
      });
      return;
    }

    if (queryParams.toLowerCase().includes("<script>")) {
      console.log("inside <Script> condition");
      setChatInputErrorMessage({
        status: true,
        message: translations.script_Query
      });
      setisResponseDone(true);
      setBrowerSearchId("");
      return;
    }

    setshowWelcomeMessage(false);
    setisResponseDone(false);

    setHome(false);
    setloadingFromRecent(false);
    setFilesToUpload([]);
    //setMessage("");
    const doc_id_list = documentStatusList
      .filter((doc) => doc.state === "uploaded")
      .map((doc) => doc.doc_id);
    setTimeout(() => {
      endOfMessageRef1.current?.scrollIntoView({ behavior: "smooth" });
      inputQueryRef.current.value = "";
    }, 200);
    setresponseValues((prev) => {
      return [
        ...prev,
        {
          content: queryParams.replaceAll("\n", "\n\n"),
          sender: "user",
          upVote: null,
          feedback: "",
          timestamp: new Date().toISOString(),
          newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
        },
        "...",
      ];
    });

    setdisplayErrorMessage({
      status: false,
      message: "",
    });
    const timeZone = Intl.DateTimeFormat()
      .resolvedOptions()
      .timeZone?.toString();
    if (chatID) {
      console.log("this is second question", responseValues.length);
      if (selectedPrompt === "rk01") {
        fnScheduleMetting(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          timeZone,
          chatID
        )
          .then((response) => {
            if (response?.status && response?.status === "error") {
              const errorMsg = response.message.llmResponse.content;

              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [
                  ...res,
                  {
                    content: errorMsg,
                    sender: "assistant",
                    status: "error",
                    timestamp: new Date().toISOString(),
                    newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                  }
                ];
              });
              setisResponseDone(true);
              return;
            }
            stream_response_handler(response, requestsession, "existing_Chat");
          })

          .catch((error) => {
            setisResponseDone(true);
            console.log("error", error);
          });
      } else if (selectedPrompt === "rk02") {
        fnWebexGroupChat(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          timeZone,
          chatID
        )
          .then((response) => {
            // need to handle the failure case status not 200 properly throw the error so it should not disturb the code flow
            if (response?.status && response?.status === "error") {
              const errorMsg = response.message.llmResponse.content;

              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [
                  ...res,
                  {
                    content: errorMsg,
                    sender: "assistant",
                    status: "error",
                    timestamp: new Date().toISOString(),
                    newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                  }
                ];
              });
              setisResponseDone(true);
              return;
            }
            stream_response_handler(response, requestsession, "existing_Chat");
          })

          .catch((error) => {
            setisResponseDone(true);
            console.log("error", error);
          });
      } else if (selectedPrompt === "rk03") {
        fnWebexMeetingSummary(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          timeZone,
          chatID
        )
          .then((response) => {
            // need to handle the failure case status not 200 properly throw the error so it should not disturb the code flow
            if (response?.status && response?.status === "error") {
              const errorMsg = response.message.llmResponse.content;

              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [
                  ...res,
                  {
                    content: errorMsg,
                    sender: "assistant",
                    status: "error",
                    timestamp: new Date().toISOString(),
                    newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                  }
                ];
              });
              setisResponseDone(true);
              return;
            }
            stream_response_handler(response, requestsession, "existing_Chat");
          })

          .catch((error) => {
            setisResponseDone(true);
            console.log("error", error);
          });
      } else if (selectedPrompt == "EL1") {
        fnDitaLangTranslation(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          chatID,
          doc_id_list
        ).then((response) => {
          if (localStorage.getItem("chat_id") !== chatID) {
             if (response && response.status !== "error") {
                setMessageReadyNotificationContext({
                  is_message_ready: true,
                  ...response,
                });
             }
             return;
          }
          if (response?.status === "error") {
            const errorMsg = response.message?.llmResponse?.content || "An error occurred";
            setresponseValues((prev) => {
              let res = [...prev];
              res.splice(-1, 1);
              return [...res, {
                content: errorMsg,
                sender: "assistant",
                status: "error",
                timestamp: new Date().toISOString(),
                newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
              }];
            });
            setisResponseDone(true);
            return;
          }

          const newMsg = {
            content: response.message?.llmResponse?.content,
            sender: response.message?.llmResponse?.sender || "assistant",
            id: response.message?.llmResponse?.message_id || response.message?.message_id,
            job_details: { job_id: null, status: "ready" },
            timestamp: response.message?.llmResponse?.timestamp,
            newTimeStamp: fnHandleTimeZone(response.message?.llmResponse?.timestamp),
            supporting_documents: response.message?.llmResponse?.message_related_document || [],
            source_url: response.source_url || [],
            upVote: null,
            message_related_document: response.message?.llmResponse?.message_related_document?.[0] || null
          };

          console.log("newMsg",newMsg)

          setresponseValues((prev) => {
            let res = [...prev];
            res.splice(-1, 1);
            return [...res, newMsg];
          });

          setisResponseDone(true);

          if (response.chat_id && response.chat_id !== chatID) {
            setchatID(response.chat_id);
            localStorage.setItem("chat_id", response.chat_id);
            fnLoadRecentChats(azureUser.accessToken);
          }

          const isNewID = response.chat_id !== chatID;
          const isFirstMessage = response.message?.llmResponse?.message_id <= 2;

          if (response.chat_id && (isNewID || isFirstMessage)) {
              fnLoadRecentChats(azureUser.accessToken);
          }
          
        }).catch((error) => {
          setisResponseDone(true);
          console.log("error", error);
        });
      } else if (selectedPrompt === "a7b2") {
        fnExcelAgentChat(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          doc_id_list,
          chatID
        )

          .then((response) => {
            if (requestsession !== localStorage.getItem("chatsessionid")) {
              if (response && response.status !== "error") {
                setMessageReadyNotificationContext({
                  is_message_ready: true,
                  ...response,
                });
              }
              fnLoadRecentChats(azureUser.accessToken);
              return; 
          }
            if (response?.status === "error") {
              const errorMsg = response.message?.llmResponse?.content || "An error occurred";
              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [...res, {
                  content: errorMsg,
                  sender: "assistant",
                  status: "error",
                  timestamp: new Date().toISOString(),
                  newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                }];
              });
              setisResponseDone(true);
              return;
            }

            console.log("print chat di here", response)
            const newMsg = {
              content: response.message?.llmResponse?.content,
              sender: response.message?.llmResponse?.sender || "assistant",
              id: response.message?.llmResponse?.message_id || response.message?.message_id,
              job_details: { job_id: null, status: "ready" },
              timestamp: response.message?.llmResponse?.timestamp,
              newTimeStamp: fnHandleTimeZone(response.message?.llmResponse?.timestamp),
              supporting_documents: response.message?.llmResponse?.supporting_documents || response.message?.llmResponse?.message_related_document || [],
              source_url: response.source_url || [],
              message_related_document: response.message?.llmResponse?.message_related_document?.[0] || null
            };

            setresponseValues((prev) => {
              let res = [...prev];
              res.splice(-1, 1);
              return [...res, newMsg];
            });

            setisResponseDone(true);

            if (response.chat_id && response.chat_id !== chatID) {
              setchatID(response.chat_id);
              localStorage.setItem("chat_id", response.chat_id);
              fnLoadRecentChats(azureUser.accessToken);
            }

            const isNewID = response.chat_id !== chatID;
            const isFirstMessage = response.message?.llmResponse?.message_id <= 2;

            if (response.chat_id && (isNewID || isFirstMessage)) {
                fnLoadRecentChats(azureUser.accessToken);
            }
          }).catch((error) => {
            setisResponseDone(true);
            inputQueryRef.current.value = null;
            console.log("error", error);
          });
      } else if(selectedPrompt === "EML"){
        fnOutlookAgent(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          chatID,
          false, // stream
        )
        .then((response) => {
            if (requestsession !== localStorage.getItem("chatsessionid")) {
              if (response && response.status !== "error") {
                setMessageReadyNotificationContext({
                  is_message_ready: true,
                  ...response,
                });
              }
              return; 
          }
            if (response?.status === "error") {
              const errorMsg = response.message?.llmResponse?.content || "An error occurred";
              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [...res, {
                  content: errorMsg,
                  sender: "assistant",
                  status: "error",
                  timestamp: new Date().toISOString(),
                  newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                }];
              });
              setisResponseDone(true);
              return;
            }

            const newMsg = {
              content: response.message?.llmResponse?.content,
              sender: response.message?.llmResponse?.sender || "assistant",
              id: response.message?.llmResponse?.message_id || response.message?.message_id,
              job_details: { job_id: null, status: "ready" },
              timestamp: response.message?.llmResponse?.timestamp,
              newTimeStamp: fnHandleTimeZone(response.message?.llmResponse?.timestamp),
              supporting_documents: response.message?.llmResponse?.message_related_document || [],
              source_url: response.source_url || [],
              upVote: null,
              message_related_document: response.message?.llmResponse?.message_related_document?.[0] || null
            };

            setresponseValues((prev) => {
              let res = [...prev];
              res.splice(-1, 1);
              return [...res, newMsg];
            });

            setisResponseDone(true);

            if (response.chat_id && response.chat_id !== chatID) {
              setchatID(response.chat_id);
              localStorage.setItem("chat_id", response.chat_id);
              fnLoadRecentChats(azureUser.accessToken);
            }
            const isNewID = response.chat_id !== chatID;
            const isFirstMessage = response.message?.llmResponse?.message_id <= 2;

            if (response.chat_id && (isNewID || isFirstMessage)) {
                fnLoadRecentChats(azureUser.accessToken);
            }

          }).catch((error) => {
            setisResponseDone(true);
            inputQueryRef.current.value = null;
            console.log("error", error);
          });

      } else if(selectedPrompt === "CON"){
        fnConfluenceAgent(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          chatID,
          false, // stream
        )
        .then((response) => {
            if (requestsession !== localStorage.getItem("chatsessionid")) {
              if (response && response.status !== "error") {
                setMessageReadyNotificationContext({
                  is_message_ready: true,
                  ...response,
                });
              }
              return; 
            }
            if (response?.status === "error") {
              const errorMsg = response.message?.llmResponse?.content || "An error occurred";
              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [...res, {
                  content: errorMsg,
                  sender: "assistant",
                  status: "error",
                  timestamp: new Date().toISOString(),
                  newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                }];
              });
              setisResponseDone(true);
              return;
            }

            const newMsg = {
              content: response.message?.llmResponse?.content,
              sender: response.message?.llmResponse?.sender || "assistant",
              id: response.message?.llmResponse?.message_id || response.message?.message_id,
              job_details: { job_id: null, status: "ready" },
              timestamp: response.message?.llmResponse?.timestamp,
              newTimeStamp: fnHandleTimeZone(response.message?.llmResponse?.timestamp),
              supporting_documents: response.message?.llmResponse?.message_related_document || [],
              source_url: response.source_url || [],
              upVote: null,
              message_related_document: response.message?.llmResponse?.message_related_document?.[0] || null
            };

            setresponseValues((prev) => {
              let res = [...prev];
              res.splice(-1, 1);
              return [...res, newMsg];
            });

            setisResponseDone(true);

            if (response.chat_id && response.chat_id !== chatID) {
              setchatID(response.chat_id);
              localStorage.setItem("chat_id", response.chat_id);
              fnLoadRecentChats(azureUser.accessToken);
            }
            
            const isNewID = response.chat_id !== chatID;
            const isFirstMessage = response.message?.llmResponse?.message_id <= 2;

            if (response.chat_id && (isNewID || isFirstMessage)) {
                fnLoadRecentChats(azureUser.accessToken);
            }

          }).catch((error) => {
            setisResponseDone(true);
            inputQueryRef.current.value = null;
            console.log("error", error);
          });
      } else {
        return new Promise((resolve, reject) => {
          fnContinueChatAPI(
            azureUser.accessToken,
            chatID,
            queryParams.replaceAll("\n", "\n\n"),
            doc_id_list,
            azureUser.selected_model.model_id,
            browerSearchId === "webs" ? browerSearchId : undefined,
            selectedPrompt==="" ? gptMode?.details?.gpt_id : selectedPrompt,
            gptMode?.details?.prompt_id ? true : false
          )
            .then((response) => {
              // need to handle the failure case status not 200 properly throw the error so it should not disturb the code flow
              if ((response?.status && response?.status === "error") || !response?.message?.llmResponse) {
                const errorMsg =
                  response?.message?.llmResponse?.content ||
                  response?.message || translations?.Response_load_error;

                setresponseValues((prev) => {
                  prev.splice(-1, 1);
                  return [
                    ...prev,
                    {
                      content: errorMsg,
                      sender: "assistant",
                      status: "error",
                      chat_id: response.chat_id ?? "",
                      timestamp: new Date().toISOString(),
                    }
                  ];
                });
                return;
              }

              let resChats = {
                content: response?.message?.llmResponse.content,
                source_url: response.message.llmResponse.source_url ?? [],
                service_used: response.message.llmResponse.service_used ?? "",
                status: response.status ?? "",
                isEmail: false,
                isMeeting: false,
                sender: "assistant",
                upVote: null,
                feedback: "",
                id: response.message.llmResponse.message_id,
                job_details: {
                  job_id: response?.message?.job_id ?? null,
                  status: "in_progress",
                  phaseTracker: response?.message?.job_id
                    ? getJobType(response?.message?.task_label)
                    : "",
                },
                chat_id: response?.chat_id,
                timestamp: response.message.llmResponse.timestamp,
                supporting_documents:
                  response?.message?.supporting_documents ?? [],
                // chunkStreams : {
                //   isLatest : true,
                //   Rendered : false,
                // },
                newTimeStamp: fnHandleTimeZone(
                  response.message.llmResponse.timestamp
                ),
              };
              setTimeout(() => {
                if (localStorage.getItem("chat_id") === response?.chat_id) {
                  if (response.message.message_id === 1) {
                    fnLoadRecentChats(azureUser.accessToken);
                  }
                  let res = [...responseValues];
                  res.splice(-1, 1);
                  res = [...res, resChats];
                  localStorage.setItem(
                    "LatestToRender",
                    Number(response.message.llmResponse.message_id)
                  );
                  setresponseValues((prev) => {
                    prev.splice(-1, 1);
                    return [...prev, resChats];
                  });
                  setTimeout(() => {
                    if (endOfMessageRef1 && endOfMessageRef1.current) {
                      endOfMessageRef1.current.scrollIntoView({
                        behavior: "smooth",
                      });
                    }
                  }, 250);

                  if (response?.message?.job_id) {
                    console.log(response);
                    //  pollExecutionPhase(localStorage.getItem('chatsessionid'), response.message.llmResponse.message_id);
                    pollDocumentJobStatus(
                      localStorage.getItem("chatsessionid"),
                      response?.chat_id,
                      response?.message?.job_id
                    )
                      .then(() => {
                        if (localStorage.getItem("chat_id") === chatID) {
                          fnGetAllChatsForChatId(azureUser.accessToken, chatID)
                            .then((res) => {
                              if (res?.status === "error") {
                                setdispError({
                                  ...dispError,
                                  status: true,
                                  message: res?.message,
                                });
                              } else {
                                console.log("fnGetAllChatsForChatId", res);
                                let resChats = [];

                                setfeedbackCategories(
                                  res?.metadata?.feedback_categories ?? []
                                );

                                res.messages.map((chat) => {
                                  if (
                                    chat?.job_details?.job_id &&
                                    chat?.job_details?.job_id ===
                                    response?.message?.job_id
                                  ) {
                                    chat.newTimeStamp = fnHandleTimeZone(
                                      chat.timestamp
                                    );
                                    resChats.push(chat);
                                  }
                                });
                                console.log("filteredreschats", resChats);
                                setisResponseDone(true);
                                setBrowerSearchId("");
                                setresponseValues((prev) =>
                                  prev.map((message) =>
                                    message?.job_details?.job_id ===
                                      resChats?.[0]?.job_details?.job_id
                                      ? resChats[0]
                                      : message
                                  )
                                );
                              }
                              setTimeout(() => {
                                if (
                                  endOfMessageRef1 &&
                                  endOfMessageRef1.current
                                ) {
                                  endOfMessageRef1.current.scrollIntoView({
                                    behavior: "smooth",
                                  });
                                }
                              }, 250);
                            })
                            .catch((error) => console.error(error));
                        } else {
                          setMessageReadyNotificationContext({
                            is_message_ready: true,
                            ...response,
                          });
                        }
                      })
                      .catch((err) => {
                        console.log(err);
                      });
                  } else {
                    setisResponseDone(true);
                    setBrowerSearchId("");
                  }
                  resolve(response);
                } else {
                  console.log(
                    "entered flash message context",
                    response?.job_details?.job_id
                  );
                  if (response.message.message_id === 1) {
                    fnLoadRecentChats(azureUser.accessToken);
                  }
                  if (response) {
                    if (response?.message?.job_id) {
                      pollDocumentJobStatus(
                        localStorage.getItem("chatsessionid"),
                        response?.chat_id,
                        response?.message?.job_id
                      )
                        .then((status) => {
                          console.log("Status", status)
                          if (localStorage.getItem("chat_id") !== chatID) {
                            setMessageReadyNotificationContext({
                              is_message_ready: true,
                              ...response,
                            });
                          }
                        })
                        .catch((err) => {
                          console.log(err);
                        });
                    } else {
                      setMessageReadyNotificationContext({
                        is_message_ready: true,
                        ...response,
                      });
                    }
                  }
                }
              }, 1000);

              if (responseValues.length === 76) {
                console.log("entered to response limit");
                setResponseValuesLimit(true);
                setdisplayErrorMessage({
                  ...displayErrorMessage,
                  status: true,
                  message: translations.message_Limit
                });
              }
            })

            .catch((error) => {
              setisResponseDone(true);
              setBrowerSearchId("");
              console.log("error", error);
              reject(error);
            }).finally(() => {
              setisResponseDone(true)
              setBrowerSearchId("");
            });
        });
      }
    } else {
      console.log("new chat");
      setisFav(false);
      console.log(selectedPrompt);
      if (selectedPrompt === "rk01") {
        fnScheduleMetting(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          timeZone
        )
          .then((response) => {
            console.log(response);
            if (response?.status && response?.status === "error") {
              const errorMsg = response.message.llmResponse.content;

              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [
                  ...res,
                  {
                    content: errorMsg,
                    sender: "assistant",
                    status: "error",
                    timestamp: new Date().toISOString(),
                    newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                  }
                ];
              });
              setisResponseDone(true);
              return;
            }
            stream_response_handler(response, requestsession, "New_Chat");
          })

          .catch((error) => {
            setisResponseDone(true);
            inputQueryRef.current.value = null;
            // setMessage("");

            console.log("error", error);
          });
      } else if (selectedPrompt === "rk02") {
        fnWebexGroupChat(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          timeZone
        )
          .then((response) => {
            if (response?.status && response?.status === "error") {
              const errorMsg = response.message.llmResponse.content;

              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [
                  ...res,
                  {
                    content: errorMsg,
                    sender: "assistant",
                    status: "error",
                    timestamp: new Date().toISOString(),
                    newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                  }
                ];
              });
              setisResponseDone(true);
              return;
            }
            stream_response_handler(response, requestsession, "New_Chat");
          })
          .catch((error) => {
            setisResponseDone(true);
            inputQueryRef.current.value = null;
            // setMessage("");

            console.log("error", error);
          });
      } else if (selectedPrompt === "rk03") {
        fnWebexMeetingSummary(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          timeZone
        )
          .then((response) => {
            if (response?.status && response?.status === "error") {
              const errorMsg = response.message.llmResponse.content;

              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [
                  ...res,
                  {
                    content: errorMsg,
                    sender: "assistant",
                    status: "error",
                    timestamp: new Date().toISOString(),
                    newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                  }
                ];
              });
              setisResponseDone(true);
              return;
            }
            stream_response_handler(response, requestsession, "New_Chat");
          })
          .catch((error) => {
            setisResponseDone(true);
            inputQueryRef.current.value = null;
            // setMessage("");

            console.log("error", error);
          });
      } else if (selectedPrompt === "EL1") {
        fnDitaLangTranslation(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          "",
          doc_id_list
        )

          .then((response) => {
            if (requestsession !== localStorage.getItem("chatsessionid")) {
              if (response && response.status !== "error") {
                setMessageReadyNotificationContext({
                  is_message_ready: true,
                  ...response,
                });
              }
              fnLoadRecentChats(azureUser.accessToken);
              return; 
          }
            if (response?.status === "error") {
              const errorMsg = response.message?.llmResponse?.content || "An error occurred";
              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [...res, {
                  content: errorMsg,
                  sender: "assistant",
                  status: "error",
                  timestamp: new Date().toISOString(),
                  newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                }];
              });
              setisResponseDone(true);
              return;
            }

            const newMsg = {
              content: response.message?.llmResponse?.content,
              sender: response.message?.llmResponse?.sender || "assistant",
              id: response.message?.llmResponse?.message_id || response.message?.message_id,
              job_details: { job_id: null, status: "ready" },
              timestamp: response.message?.llmResponse?.timestamp,
              newTimeStamp: fnHandleTimeZone(response.message?.llmResponse?.timestamp),
              supporting_documents: response.message?.llmResponse?.message_related_document || [],
              source_url: response.source_url || [],
              upVote: null,
              message_related_document: response.message?.llmResponse?.message_related_document?.[0] || null
            };

            setresponseValues((prev) => {
              let res = [...prev];
              res.splice(-1, 1);
              return [...res, newMsg];
            });

            setisResponseDone(true);

            if (response.chat_id && response.chat_id !== chatID) {
              setchatID(response.chat_id);
              localStorage.setItem("chat_id", response.chat_id);
              fnLoadRecentChats(azureUser.accessToken);
            }
            const isNewID = response.chat_id !== chatID;
            const isFirstMessage = response.message?.llmResponse?.message_id <= 2;

            if (response.chat_id && (isNewID || isFirstMessage)) {
                fnLoadRecentChats(azureUser.accessToken);
            }

          }).catch((error) => {
            setisResponseDone(true);
            inputQueryRef.current.value = null;
            console.log("error", error);
          });
      } else if (selectedPrompt === "a7b2") {
        fnExcelAgentChat(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          doc_id_list
        )

          .then((response) => {
            if (requestsession !== localStorage.getItem("chatsessionid")) {
              if (response && response.status !== "error") {
                setMessageReadyNotificationContext({
                  is_message_ready: true,
                  ...response,
                });
              }
              fnLoadRecentChats(azureUser.accessToken);
              return; 
          }
            if (response?.status === "error") {
              const errorMsg = response.message?.llmResponse?.content || "An error occurred";
              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [...res, {
                  content: errorMsg,
                  sender: "assistant",
                  status: "error",
                  timestamp: new Date().toISOString(),
                  newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                }];
              });
              setisResponseDone(true);
              return;
            }

            const newMsg = {
              content: response.message?.llmResponse?.content,
              sender: response.message?.llmResponse?.sender || "assistant",
              id: response.message?.llmResponse?.message_id || response.message?.message_id,
              job_details: { job_id: null, status: "ready" },
              timestamp: response.message?.llmResponse?.timestamp,
              newTimeStamp: fnHandleTimeZone(response.message?.llmResponse?.timestamp),
              supporting_documents: response.message?.llmResponse?.supporting_documents || response.message?.llmResponse?.message_related_document || [],
              source_url: response.source_url || [],
              message_related_document: response.message?.llmResponse?.message_related_document?.[0] || null
            };

            console.log("print chat di here", response)

            setresponseValues((prev) => {
              let res = [...prev];
              res.splice(-1, 1);
              return [...res, newMsg];
            });

            setisResponseDone(true);

            if (response.chat_id && response.chat_id !== chatID) {
              setchatID(response.chat_id);
              localStorage.setItem("chat_id", response.chat_id);
              fnLoadRecentChats(azureUser.accessToken);
            }

            const isNewID = response.chat_id !== chatID;
            const isFirstMessage = response.message?.llmResponse?.message_id <= 2;

            if (response.chat_id && (isNewID || isFirstMessage)) {
                fnLoadRecentChats(azureUser.accessToken);
            }
          }).catch((error) => {
            setisResponseDone(true);
            inputQueryRef.current.value = null;
            console.log("error", error);
          });
      } else if(selectedPrompt === "EML"){
        fnOutlookAgent(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          null, // chatId
          false // stream
        )
          .then((response) => {
            if(requestsession !== localStorage.getItem("chatsessionid")){
              if(response && response.status !== "error") {
                setMessageReadyNotificationContext({
                  is_message_ready: true,
                  ...response,
                })
              }
              fnLoadRecentChats(azureUser.accessToken);
              return;
            }
            if (response?.status === "error") {
              const errorMsg = response.message?.llmResponse?.content || "An error occurred";
              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [...res, {
                  content: errorMsg,
                  sender: "assistant",
                  status: "error",
                  timestamp: new Date().toISOString(),
                  newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                }];
              });
              setisResponseDone(true);
              return;
            }
            const newMsg = {
              content: response.message?.llmResponse?.content,
              sender: response.message?.llmResponse?.sender || "assistant",
              id: response.message?.llmResponse?.message_id || response.message?.message_id,
              job_details: { job_id: null, status: "ready" },
              timestamp: response.message?.llmResponse?.timestamp,
              newTimeStamp: fnHandleTimeZone(response.message?.llmResponse?.timestamp),
              supporting_documents: response.message?.llmResponse?.supporting_documents || response.message?.llmResponse?.message_related_document || [],
              source_url: response.source_url || [],
              message_related_document: response.message?.llmResponse?.message_related_document?.[0] || null
            };

            setresponseValues((prev) => {
              let res = [...prev];
              res.splice(-1, 1);
              return [...res, newMsg];
            });

            setisResponseDone(true);

            if (response.chat_id && response.chat_id !== chatID) {
              setchatID(response.chat_id);
              localStorage.setItem("chat_id", response.chat_id);
              fnLoadRecentChats(azureUser.accessToken);
            }

            const isNewID = response.chat_id !== chatID;
            const isFirstMessage = response.message?.llmResponse?.message_id <= 2;

            if (response.chat_id && (isNewID || isFirstMessage)) {
                fnLoadRecentChats(azureUser.accessToken);
            }
          }).catch((error) => {
            setisResponseDone(true);
            inputQueryRef.current.value = null;
            console.log("error", error);
          });
      } else if(selectedPrompt === "CON"){
        fnConfluenceAgent(
          azureUser.accessToken,
          queryParams.replaceAll("\n", "\n\n"),
          null, // chatId
          false // stream
        ).then((response) => {
            if(requestsession !== localStorage.getItem("chatsessionid")){
              if(response && response.status !== "error") {
                setMessageReadyNotificationContext({
                  is_message_ready: true,
                  ...response,
                })
              }
              fnLoadRecentChats(azureUser.accessToken);
              return;
            }

            if (response?.status === "error") {
              const errorMsg = response.message?.llmResponse?.content || "An error occurred";
              setresponseValues((prev) => {
                let res = [...prev];
                res.splice(-1, 1);
                return [...res, {
                  content: errorMsg,
                  sender: "assistant",
                  status: "error",
                  timestamp: new Date().toISOString(),
                  newTimeStamp: fnHandleTimeZone(new Date().toISOString()),
                }];
              });
              setisResponseDone(true);
              return;
            }
            
            const newMsg = {
              content: response.message?.llmResponse?.content,
              sender: response.message?.llmResponse?.sender || "assistant",
              id: response.message?.llmResponse?.message_id || response.message?.message_id,
              job_details: { job_id: null, status: "ready" },
              timestamp: response.message?.llmResponse?.timestamp,
              newTimeStamp: fnHandleTimeZone(response.message?.llmResponse?.timestamp),
              supporting_documents: response.message?.llmResponse?.supporting_documents || response.message?.llmResponse?.message_related_document || [],
              source_url: response.source_url || [],
              message_related_document: response.message?.llmResponse?.message_related_document?.[0] || null
            };
 
            setresponseValues((prev) => { 
              let res = [...prev]; 
              res.splice(-1, 1); 
              return [...res, newMsg]; 
            });

            setisResponseDone(true);
            
            if (response.chat_id && response.chat_id !== chatID) {
              setchatID(response.chat_id);
              localStorage.setItem("chat_id", response.chat_id);
              fnLoadRecentChats(azureUser.accessToken);
            }

            const isNewID = response.chat_id !== chatID;
            const isFirstMessage = response.message?.llmResponse?.message_id <= 2;

            if (response.chat_id && (isNewID || isFirstMessage)) {
                fnLoadRecentChats(azureUser.accessToken);
            }
          }).catch ((error) => {
            setisResponseDone(true);
            inputQueryRef.current.value = null;
            console.log("error", error);
          });
      } else {
        const default_model_for_prompt =
          selectedPrompt !== ""
            ? azureUser?.models_accessible?.filter(
              (model) => model.llm_name === "GPT-4o"
            )?.[0]?.model_id
            : azureUser.selected_model.model_id;

        let default_prompt = selectedPrompt;
        if (selectedPrompt === "" && browerSearchId === "webs") {
          default_prompt = browerSearchId;
        }
        return new Promise((resolve, reject) => {
          console.log("-------GPT Detai;ls,", gptMode?.details)
          fnNewChatAPI(
            azureUser.accessToken,
            queryParams.replaceAll("\n", "\n\n"),
            doc_id_list,
            default_prompt,
            azureUser.selected_model.model_id,
            folder_id,
            gptMode?.details?.prompt_id ? true : false
          )
            .then((response) => {
              let resChats = {
                content: response?.message?.llmResponse.content,
                source_url: response.message.llmResponse.source_url ?? [],
                service_used: response.message.llmResponse.service_used ?? "",
                status: response.status ?? "",
                isEmail: false,
                isMeeting: false,
                sender: "assistant",
                upVote: null,
                feedback: "",
                id: response.message.llmResponse.message_id,
                job_details: {
                  job_id: response?.message?.job_id ?? null,
                  status: "in_progress",
                  phaseTracker: response?.message?.job_id
                    ? getJobType(response?.message?.task_label)
                    : "",
                },
                chat_id: response.chat_id,
                supporting_documents:
                  response?.message?.supporting_documents ?? [],
                timestamp: response.message.llmResponse.timestamp,
                newTimeStamp: fnHandleTimeZone(
                  response.message.llmResponse.timestamp
                ),
              };
              console.log(localStorage?.getItem("chat_id"), "chat id ");
              if (localStorage?.getItem("chat_id")) {
                if (localStorage.getItem("chat_id") !== response.chat_id) {
                  console.log(
                    "return form new api, 'newchat->existing",
                    localStorage.getItem("chat_id")
                  );
                  console.log("entered flash message context");
                  if (response?.message?.job_id) {
                    pollDocumentJobStatus(
                      localStorage.getItem("chatsessionid"),
                      response?.chat_id,
                      response?.message?.job_id
                    )
                      .then((status) => {
                        console.log("status", status)
                        if (localStorage.getItem("chat_id") !== chatID) {
                          setMessageReadyNotificationContext({
                            is_message_ready: true,
                            ...response,
                          });
                        }
                      })
                      .catch((err) => {
                        console.log(err);
                      });
                  } else {
                    if (response)
                      setMessageReadyNotificationContext({
                        is_message_ready: true,
                        ...response,
                      });
                  }
                  fnLoadRecentChats(azureUser.accessToken);
                  return;
                }
              } else {
                //  setfileName(response.fileName);
                // setMessage("");
                localStorage.setItem(
                  "LatestToRender",
                  Number(response.message.llmResponse.message_id)
                );
                setfeedbackCategories(
                  response.metadata?.feedback_categories ?? []
                );
                setTimeout(() => {
                  console.log(responseValues);
                  //!localStorage.getItem("chat_id")
                  if (
                    localStorage.getItem("chat_id") === response.chat_id ||
                    requestsession === localStorage.getItem("chatsessionid")
                  ) {
                    let res = [...responseValues];
                    res.splice(-1, 1);
                    res = [...res, resChats];
                    setresponseValues((prev) => {
                      prev.splice(-1, 1);
                      return [...prev, resChats];
                    });
                    setTimeout(() => {
                      if (endOfMessageRef1 && endOfMessageRef1.current) {
                        endOfMessageRef1.current.scrollIntoView({
                          behavior: "smooth",
                        });
                      }
                    }, 250);
                    setTimeout(
                      () => localStorage.setItem("LatestToRender", ""),
                      200
                    );
                    setchatID(response.chat_id);
                    localStorage.setItem("chat_id", response.chat_id);

                    if (response?.message?.job_id) {
                      // pollExecutionPhase(localStorage.getItem('chatsessionid'), response.message.llmResponse.message_id);
                      pollDocumentJobStatus(
                        localStorage.getItem("chatsessionid"),
                        response.chat_id,
                        response?.message?.job_id
                      )
                        .then(() => {
                          if (
                            localStorage.getItem("chat_id") === chatID ||
                            !chatID
                          ) {
                            fnGetAllChatsForChatId(
                              azureUser.accessToken,
                              response.chat_id
                            )
                              .then((res) => {
                                if (res?.status === "error") {
                                  setdispError({
                                    ...dispError,
                                    status: true,
                                    message: res?.message,
                                  });
                                } else {
                                  localStorage.setItem(
                                    "LatestToRender",
                                    Number(
                                      response.message.llmResponse.message_id
                                    )
                                  );
                                  console.log("fnGetAllChatsForChatId", res);
                                  let resChats = [];

                                  setfeedbackCategories(
                                    res?.metadata?.feedback_categories ?? []
                                  );

                                  res.messages.map((chat) => {
                                    if (
                                      chat?.job_details?.job_id &&
                                      chat?.job_details?.job_id ===
                                      response?.message?.job_id
                                    ) {
                                      chat.newTimeStamp = fnHandleTimeZone(
                                        chat.timestamp
                                      );
                                      resChats.push(chat);
                                    }
                                  });
                                  console.log("filteredreschats", resChats);
                                  setisResponseDone(true);
                                  setBrowerSearchId("");
                                  setresponseValues((prev) =>
                                    prev.map((message) =>
                                      message?.job_details?.job_id ===
                                        resChats?.[0]?.job_details?.job_id
                                        ? resChats[0]
                                        : message
                                    )
                                  );
                                  setTimeout(() => {
                                    if (
                                      endOfMessageRef1 &&
                                      endOfMessageRef1.current
                                    ) {
                                      endOfMessageRef1.current.scrollIntoView({
                                        behavior: "smooth",
                                      });
                                    }
                                  }, 250);
                                }
                              })
                              .catch((error) => console.error(error));
                          } else {
                            setMessageReadyNotificationContext({
                              is_message_ready: true,
                              ...response,
                            });
                          }
                        })
                        .catch((err) => {
                          console.log(err);
                        });
                    } else {
                      setisResponseDone(true);
                      setBrowerSearchId("");
                    }
                    resolve(response);
                  } else {
                    console.log(
                      "entered flash message context",
                      response?.job_details?.job_id
                    );
                    if (response) {
                      console.log(response);
                      if (response?.message?.job_id) {
                        pollDocumentJobStatus(
                          localStorage.getItem("chatsessionid"),
                          response?.chat_id,
                          response?.message?.job_id
                        )
                          .then(() => {
                            if (localStorage.getItem("chat_id") !== chatID) {
                              setMessageReadyNotificationContext({
                                is_message_ready: true,
                                ...response,
                              });
                            }
                          })
                          .catch((err) => {
                            console.log(err);
                          });
                      } else {
                        setMessageReadyNotificationContext({
                          is_message_ready: true,
                          ...response,
                        });
                      }
                    }
                  }
                }, 1000);
                fnLoadRecentChats(azureUser.accessToken);
              }
            })
            .catch((error) => {
              setisResponseDone(true);
              setBrowerSearchId("");
              inputQueryRef.current.value = null;
              // setMessage("");
              reject(error);
              console.log("error", error);
            });
        });
      }
    }
    setTimeout(() => {
      const textarea = document.getElementById("inputQuery");
      textarea.focus();
    }, 100);
  }

  const fnLoadChat = (
    e,
    chat_id,
    is_favourite,
    document_id,
    fileList,
    model_id,
    llm_name,
    chat_block_context,
    folder_id = null
  ) => {
    localStorage.setItem("chat_id", chat_id);

    localStorage.setItem(
      "chatsessionid",
      Number(localStorage.getItem("chatsessionid")) + 1
    );
    if (inputQueryRef && inputQueryRef?.current) {
      inputQueryRef.current.value = "";
      inputQueryRef.current.focus();
    }
    // setMessage("");
    resetInlineVoiceChat();
    // setTimeout(() => {
    //   const textarea = document.getElementById("inputQuery");
    //   if (textarea) textarea.style.height = "60px";
    // }, 500);

    const defaultmodelForchat =
      azureUser?.models_accessible?.filter(
        (model) => model.model_id === model_id
      )?.[0] ?? [];

    setazureUser({
      ...azureUser,
      selected_model: {
        model_id: defaultmodelForchat?.model_id ?? model_id,
        model_name: defaultmodelForchat?.llm_name ?? llm_name,
        document_supported: defaultmodelForchat?.document_supported ?? false,
        default_model: defaultmodelForchat?.default_model ?? false,
      },
    });
    setChatInputErrorMessage({ status: false, message: "" });
    setIsFileProcessing(false);
    setselectedPrompt("");
    setchatID(chat_id);
    setFolder_id(folder_id);
    setblnLeftNavigationClicked(true);
    setisResponseDone(true);
    setOpenSearch(false);
    setHome(false);
    setloadingFromRecent(true);
    setshowWelcomeMessage(false);
    setResponseValuesLimit(false);
    // setfileName(e.target.getAttribute("chat_file_name"));
    // setdocument_id(e.target.getAttribute("document_id"));
    setdocument_id(document_id);
    setBlocked_chat_context({
      is_blocked: chat_block_context.is_blocked,
      blocked_message: chat_block_context.blocked_message,
    });
    setFileList(fileList);
    setDocumentStatusList([]);
    setFilesToUpload([]);
    setFileUploadComponentErros({ type: "", message: "", status: false });
    setisFav(is_favourite === "true" ? true : false);
    if (displayErrorMessage.status)
      setdisplayErrorMessage({
        ...displayErrorMessage,
        status: false,
        message: "",
      });
    setBrowerSearchId("");
    fnGetAllChatsForChatId(azureUser.accessToken, chat_id)
      .then((res) => {
        if (res?.status === "error") {
          setdispError({
            ...dispError,
            status: true,
            message: res?.message,
          });

          setchatID();
          localStorage.setItem("chat_id", "");
          setresponseValues([]);
          setdocument_id([]);
          setFileList([]);
          setshowWelcomeMessage(false);
          setHome(true);
          setloadingFromRecent(false);
        } else {
          const isActualEL1Task = res.messages?.some((msg) =>
            msg.sender === "assistant" &&
            msg.content?.toLowerCase().includes("upload your xml documents")
          )
          
          if (isActualEL1Task) {
            setselectedPrompt("EL1");
          } else {
            setselectedPrompt("");
          }
          let resChats = [];

          let fileToUploadViewList = [...res?.documents];
          setDocumentStatusList(
            res?.documents
              ?.filter((doc) => doc.document_status === "Success")
              ?.map((doc) => {
                return {
                  id: doc.original_file_name,
                  valid: true,
                  status: doc.document_status,
                  message: "",
                  doc_id: doc.document_id,
                  message_id: doc.message_id - 1,
                  state: "uploaded",
                  key: crypto.randomUUID(),
                };
              })
          );

          setfeedbackCategories(res?.metadata?.feedback_categories ?? []);
          localStorage.setItem("LatestToRender", "");
          res.messages.map((chat) => {
            chat.newTimeStamp = fnHandleTimeZone(chat.timestamp);
            if (chat.job_details?.job_id) {
              chat.job_details.phaseTracker = getJobType(
                chat.job_details.task_label
              );
            }
            resChats.push(chat);
            if (
              chat.job_details?.job_id &&
              chat.job_details?.status === "in_progress"
            ) {
              console.log("polling initiated at get chat messages");
              pollDocumentJobStatus(
                localStorage.getItem("chatsessionid"),
                chat_id,
                chat.job_details?.job_id
              )
                .then((xx) => {
                  console.log(xx);
                  console.log("eneted load trigger ");
                  if (localStorage.getItem("chat_id") === chat_id) {
                    fnGetAllChatsForChatId(azureUser.accessToken, chat_id)
                      .then((res) => {
                        if (res?.status === "error") {
                          setdispError({
                            ...dispError,
                            status: true,
                            message: res?.message,
                          });
                        } else {
                          console.log("fnGetAllChatsForChatId", res);
                          let resChats = [];
                          setfeedbackCategories(
                            res?.metadata?.feedback_categories ?? []
                          );
                          res.messages.map((loadchat) => {
                            if (
                              loadchat?.job_details?.job_id &&
                              loadchat?.job_details?.job_id ===
                              chat?.job_details?.job_id
                            ) {
                              loadchat.newTimeStamp = fnHandleTimeZone(
                                loadchat.timestamp
                              );
                              resChats.push(loadchat);
                            }
                          });
                          console.log("filteredreschats", resChats);
                          setresponseValues((prev) =>
                            prev.map((message) =>
                              message?.job_details?.job_id ===
                                resChats?.[0]?.job_details?.job_id
                                ? resChats?.[0]
                                : message
                            )
                          );
                        }
                      })
                      .catch((error) => console.error(error));
                  }
                })
                .catch((err) => {
                  console.log(err);
                });
            }

            fileToUploadViewList = res?.documents.filter((doc) => {
              if (
                doc.message_id > chat.id &&
                doc.document_status === "Success"
              ) {
                return doc;
              }
            });
          });

          setFilesToUpload(
            fileToUploadViewList?.map((doc) => {
              return {
                id: doc.original_file_name,
                valid: true,
                status: doc.document_status,
                message: "",
                doc_id: doc.document_id,
                message_id: doc.message_id - 1,
                state: "uploaded",
                key: crypto.randomUUID(),
              };
            })
          );

          if (resChats.length === 76) {
            setResponseValuesLimit(true);
            setdisplayErrorMessage({
              ...displayErrorMessage,
              status: true,
              message: translations.message_Limit
            });
          }

          if (resChats[resChats.length - 1].sender === "user") {
            resChats = [...resChats, "..."];
          }

          setresponseValues(() => {
            return resChats;
          });

          setTimeout(() => {
            endOfMessageRef1.current?.scrollIntoView({ behavior: "instant" });
            if (inputQueryRef && inputQueryRef.current)
              inputQueryRef.current.value = "";
          }, 200);

          setloadingFromRecent(false);
        }
      })
      .catch((error) => console.error(error));
  };

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const fnHandleFilterRecentchats = () => {
    // setresChatSearch(e.detail);
    let actualchats = recentChats ?? [];
    let search_term = resChatSearch?.toLowerCase().trim();
    if (segmentControlState === 0) {
      let filtered = [...actualchats];

      filtered = favRecentChats.filter(
        (item) =>
          item.isFavorite &&
          item?.title.toLowerCase().trim().includes(search_term)
      );
      setFilteredFavRecentChats(filtered);
      dispatch(setRecentChats(filtered));
    } else {
      let filtered = [...actualchats];
      filtered = actualchats.filter((item) =>
        item?.title.toLowerCase().trim().includes(search_term)
      );
      setFilteredRecentChats(filtered);
      dispatch(setRecentChats(filtered));
    }
  };

  const fnMarkRecentGpt = (accessToken, gptId) => {
    fnRecentGPT(accessToken, gptId)
      .then((response) => {
        console.log("Success in fnMarkRecentGpt", response)
      })
      .catch((error) => {
        console.error("Error marking GPT as recent:", error);
        setdispAlertMessage({
          status: true,
          code: "error",
          message: "Failed to mark GPT as recent. Please try again.",
        });
      });
  }

const fnClickPrompt = async (
    e,
    title,
    prompt_id,
    promt_mode,
    requestsession,
    is_gpt = false,
    gpt_data = null
  ) => {
    fnMarkRecentGpt(azureUser?.accessToken, gpt_data?.gpt_id)

    let newArray = promptsData.filter((val) => val.prompt_id === prompt_id);

    localStorage.setItem("isGptPrompt", true)
    if (newArray.length > 0) {
      setselectedPrompt(newArray?.[0]?.["prompt_id"]);
    }
    if (["rk01", "rk02", "rk03"].includes(newArray?.[0]?.prompt_id)) {
      setisLoading(true);
      const result = await fnGetMeetingSchedulerAgentDetails(
        azureUser.accessToken,
        translations?.Connectivity_issue
      );
      setisLoading(false);
      if (!result || result.status !== 200) {
        console.error(result?.message || translations.agent_Detail_Fail);
        setdispAlertMessage({
          status: true,
          code: "error",
          message: result?.message || translations.agent_Detail_Fail,
        });

        return;
      }

      if (result.response && !result.response.isAuthenticated) {
        window.open(result.response.auth_url, "_blank");
        return;
      }
    }

    if (openGPTScreen) {
      setOpenGPTScreen(!openGPTScreen);
    }
    if (newOpenGPTScreen) {
      setNewOpenGPTScreen(!newOpenGPTScreen)
    }
    resetInlineVoiceChat();

    setloadingFromRecent(true);
    setHome(false);
    setshowWelcomeMessage(false);
    setisPromptClicked(false);
    setchatID();
    localStorage.setItem("chat_id", "");
    // setfileName("");
    setFileList([]);
    setdocument_id([]);
    setBlocked_chat_context({
      is_blocked: false,
      blocked_message: "",
    });
    setDocumentStatusList([]);
    setisFav(false);
    setFilesToUpload([]);
    setFileUploadComponentErros({ type: "", message: "", status: false });

    console.log("newArray", newArray);
    if(prompt_id === "GV1"){
      console.log("GV1 prompt clicked")
      setloadingFromRecent(false);
      setOpenAIGIFlow(true);
      return;
    }
     if (prompt_id === "IJI") {
    setOpenFIJ(true);
      }
    if (prompt_id === "BPQ") {
      setOpenBPQ(true);
      setloadingFromRecent(false);
      return;
    }
    else{

    
    if (
      newArray.length > 0 &&
      (newArray?.[0]?.["prompt_mode"] === "general" ||
        newArray?.[0]?.["prompt_mode"] === "document_upload")
    ) {
      if (localStorage.getItem("chatsessionid") === requestsession) {
        setresponseValues(() => {
          return [
            {
              // Instead of content, store the KEY
              contentKey: `${newArray[0].prompt_id}_user_utterance`,
              isTranslatable: true, // Add a flag
              sender: "user",
              upVote: null,
              feedback: "",
            },
            "...",
          ];
        });

        setloadingFromRecent(false);

        if (newArray?.[0]?.["prompt_id"] === "ekm4") {
          const default_model = (azureUser.models_accessible ?? []).filter(
            (model) => model.default_model
          );
          setazureUser((prev) => {
            return {
              ...prev,
              selected_model: {
                model_id: default_model?.[0]?.model_id,
                model_name: default_model?.[0]?.llm_name,
                document_supported: default_model?.[0]?.document_supported,
                default_model: default_model?.[0]?.default_model,
              },
            };
          });
        }
        setTimeout(() => {
          setresponseValues((prev) => {
            prev.splice(-1, 1);
            if (localStorage.getItem("chatsessionid") === requestsession)
              return [
                ...prev,
                {
                  contentKey: `${newArray[0].prompt_id}_bot_utterance`,
                  isTranslatable: true,
                  sender: "assistant",
                  upVote: null,
                  feedback: "",
                  id: "",
                },
              ];
            else {
              return [];
            }
          });
        }, 2000);

        setTimeout(() => {
          endOfMessageRef1?.current?.scrollIntoView({ behavior: "smooth" });
          if (inputQueryRef && inputQueryRef.current)
            inputQueryRef.current.value = "";
        }, 200);

        setGptMode({ is_gpt: is_gpt, details: gpt_data });
      }

    } else if (newArray?.[0]?.["prompt_mode"] === "Webex_Agent") {
      if (localStorage.getItem("chatsessionid") === requestsession) {
        setresponseValues([
          {
            content: newArray[0]["prompt_bot_utterance"],
            sender: "assistant",
            upVote: null,
            feedback: "",
            id: "",
          },
        ]);

        setloadingFromRecent(false);

        if (["rk02", "rk03"].includes(newArray?.[0]?.["prompt_id"])) {
          console.log("in the webex agent", newArray?.[0]?.["prompt_id"]);
          setTimeout(
            () =>
              fnHandleSubmit(
                localStorage.getItem("chatsessionid"),
                newArray?.[0]?.["prompt_id"],
                "Webex_Groups_Session",
                newArray?.[0]?.["prompt_user_utterance"]
              ),
            500
          );
        }
      }
    } else if (newArray?.[0]?.["prompt_mode"] === "user_defined" || promt_mode === "user_defined") {

      if (["max3"].includes(newArray?.[0]?.["prompt_id"])) {
        setresponseValues([
          {
            content: newArray[0]["prompt_bot_utterance"],
            sender: "assistant",
            upVote: null,
            feedback: "",
            id: "",
          },
        ]);
        setshowWelcomeMessage(false);
      } else {
        setresponseValues([]);
        setshowWelcomeMessage(true);
      }

      setselectedPrompt(newArray?.[0]?.["prompt_id"] ?? prompt_id);
      console.log("---------- is_gpt", is_gpt)
      if (is_gpt) setGptMode({ is_gpt: is_gpt, details: gpt_data });
      setTimeout(() => setloadingFromRecent(false), 500);
    }
  }
    const model = azureUser?.models_accessible?.filter((item) => item?.model_id === gpt_data?.model)

    if (model[0]) {
      setazureUser((prev) => {
        return {
          ...prev,
          selected_model: {
            model_id: model[0]?.model_id,
            model_name: model[0]?.llm_name,
            document_supported: model[0]?.document_supported,
            default_model: model[0]?.default_model,
          },
        };
      });
    }

      const msg = {
        EL1: `${translations?.Xml_welcome_msg}\n\n${translations?.Xml_upload_instruction}\n\n${translations?.Xml_help_msg}`,
        EML: translations?.eml_welcome_msg,
        CON: translations?.con_welcome_msg,
      };

      if (msg[prompt_id]) {
        setresponseValues([
          {
            content: msg[prompt_id],
            sender: "assistant",
            upVote: null,
            feedback: "",
            id: "",
          },
        ]);
      }

  };



  function segreateFolders(folders) {
    const fav_folders = [];
    folders.forEach((folder) => {
      if (Array.isArray(folder.children) && folder.children.length > 0) {
        console.log(folder);
        let children = [];
        folder.children.forEach((child) => {
          if (child.isFavorite) {
            children.push(child);
          }
        });
        if (children.length > 0)
          fav_folders.push({
            label: folder.label,
            value: folder.value,
            children: children,
          });
      }
    });
    console.log(fav_folders);
    return fav_folders;
  }

  const fnLoadRecentChats = (tokenValue) => {
    setChatHistoryLoading(true);
    fnRecentChats(tokenValue).then((response) => {
      if (response === "error" || response === null || response === undefined) {
        setChatHistoryLoading(false);
        setdispAlertMessage({
          status: true,
          code: "error",
          message: translations.recent_Load_Chat_Failed_Message
        });
        return;
      }
      console.log("-----------------", response);
      let resChats = [];
      let favreschats = [];
      let folder = [];

      allchatsRec.current = response?.chats || [];

      response?.folders?.forEach((f) => {
        folder.push({ label: f.folder_name, value: f.folder_id, children: [] });
      });

      response?.chats?.map((chat) => {
        if (chat?.folder_id) {
          let folder_id = chat.folder_id;
          console.log(folder_id);
          let x = response.folders.find((f) => f.folder_id === folder_id);
          console.log(x?.folder_id, chat, response.folders);
          if (x) {
            folder = folder.map((folder) => {
              if (folder.value === folder_id) {
                let assignedChildren = folder?.children ?? [];
                folder.children = [
                  ...assignedChildren,
                  {
                    label: chat.title,
                    value: chat.id,
                    children: null,
                    ...chat,
                  },
                ];
              }
              return folder;
            });
          }
        } else {
          resChats.push(chat);
          if (chat.isFavorite === true) {
            favreschats.push(chat);
          }
        }
      });

      console.log(folder);
      dispatch(setRecentChats(resChats));
      const fav_folders = segreateFolders(folder);

      if (localStorage.getItem("LatestToRender") > 0)
        localStorage.setItem("LatestToRender", "");
      setAllFolders(folder);
      setrecentChats(() => {
        return resChats;
      });
      setFilteredRecentChats(() => {
        return resChats;
      });
      setFavRecentChats(favreschats);
      setFilteredFavRecentChats(favreschats);
      setresChatSearch("");
      setChatHistoryLoading(false);
    });
  };

  const fnSendLikeFeedback = (e, id) => {
    console.log(id);
    let msgid = String(id); // converting id into string so that item.id + "" === msgid this comparision could work

    setresponseValues((prev) => {
      return prev.map((item) => {
        if (item.id + "" === msgid) {
          return { ...item, upVote: true };
        } else {
          return item;
        }
      });
    });
    fnSendFeedbackResponse(azureUser.accessToken, chatID, msgid, true, "", "");
  };

  const fnsendDislikeFeedback = (messageid, message, messagetext) => {
    console.log(messageid, message[0]);
    console.log(messagetext);

    fnSendFeedbackResponse(
      azureUser.accessToken,
      chatID,
      messageid,
      false,
      message,
      messagetext
    ).then((resp) => {
      if (resp.message === translations.send_Feedback) {
        setresponseValues((prev) => {
          return prev.map((item) => {
            if (item.id === messageid) {
              return {
                ...item,
                upVote: false,
                feedback_categories: message,
                feedback: messagetext,
              };
            } else {
              return item;
            }
          });
        });
      }
      console.log("first", resp);
      console.log("first", resp.statuscode);
    });
  };

  const updateStatus = (message_id, documentId, status, message, state) => {
    console.log("updated status");
    setDocumentStatusList((prev) =>
      prev.map((item) =>
        item.doc_id === documentId
          ? { ...item, status, message, message_id, state }
          : item
      )
    );
    setFilesToUpload((prev) =>
      prev.map((item) =>
        item.doc_id === documentId
          ? { ...item, status, message, message_id, state }
          : item
      )
    );
  };

  const fnHandleMultiUpload = (e, chatId, requestsession) => {
    console.log("chatid", chatId);
    console.log("localstorage", localStorage.getItem("chat_id"));
    if (localStorage.getItem("chat_id") === chatId) {
      console.log(true);
    } else {
      console.log(false);
    }
    setisResponseDone(false);
    setIsFileProcessing(true);
    setisLoading(true);
    // setFileUploadStatus([]);
    // setFileUploadSuccessList([]);
    setTimeout(() => {
      setisLoading(false);
      setuploadHanledClick(!uploadHanledClick);
      if (isHome) {
        setHome(false);
      }
      if (selectedPrompt.length === 0) setshowWelcomeMessage(true);
    }, 1000);
    const filesToUpload = [...e.detail];
    const successList = [];
    const errorList = [];
    // const updatedStatusList = [];
    const documentIdList = [];

    if (localStorage.getItem("chat_id") === chatId) {
      if (document_id.length > 0 && documentStatusList.length === 0) {
        document_id.forEach((doc) => {
          setDocumentStatusList((prev) => [
            ...prev,
            { id: doc, status: "Success", message: "" },
          ]);
        });
      }
    }
    const updateStatus = (filename, documentId, status, message) => {
      setDocumentStatusList((prev) =>
        prev.map((item) =>
          item.id === filename && item.doc_id === documentId
            ? { ...item, status, message }
            : item
        )
      );
    };
    if (
      localStorage.getItem("chat_id") === chatId ||
      localStorage.getItem("chatsessionid") === requestsession
    ) {
      const uploadPromises = filesToUpload.map((file) => {
        let fileName = file.name;
        return fnGetPresingedURL(azureUser.accessToken, fileName, chatId)
          .then((response) => {
            console.log(response.status);
            const { presigned_url, document_id, file_name_full } = response;
            // updatedStatusList.push({id : document_id, status: "In Progress"});
            let fileobj = file;
            let filenameHeader = fileName;
            if (file_name_full !== fileName) {
              filenameHeader = file_name_full;
              fileobj = new File([file], file_name_full, { type: file.type });
            }
            console.log(fileobj);
            if (
              localStorage.getItem("chat_id") === chatId ||
              localStorage.getItem("chatsessionid") === requestsession
            ) {
              setDocumentStatusList((prev) => [
                ...prev,
                {
                  id: fileName,
                  status: "In Progress",
                  message: "",
                  doc_id: document_id,
                  type: file.type,
                },
              ]);
            }
            const myHeaders = new Headers();
            myHeaders.append("Content-Type", fileobj.type);
            myHeaders.append("x-amz-meta-filename", filenameHeader);
            if (chatID) myHeaders.append("x-amz-meta-chat_id", chatId);

            return fetch(presigned_url, {
              method: "PUT",
              body: fileobj,
              headers: myHeaders,
            }).then((uploadResponse) => {
              const statuscode = uploadResponse.status;
              if (!uploadResponse.ok) {
                if (statuscode === 400) {
                  if (
                    uploadResponse.message ===
                    translations.extensionUnsupported
                  ) {
                    throw new Error(`${uploadResponse.message}`);
                  } else {
                    throw new Error("File upload failed. Please try again.");
                  }
                } else if (statuscode === 403)
                  throw new Error("Unable to verify token. Please try again.");
                else throw new Error("File upload failed. Please try again.");
              }
              return document_id;
            });
          })
          .then((documentid) => {
            return pollDocumentStatus(documentid)
              .then((res) => {
                console.log(res);
                if (
                  localStorage.getItem("chat_id") === chatId ||
                  localStorage.getItem("chatsessionid") === requestsession
                ) {
                  setFileList((prev) => [...prev, fileName]);
                  //  setdocument_id((prev) => [...prev, documentid]);
                  documentIdList.push(documentid);
                  successList.push({
                    filename: fileName,
                    status: "success",
                    message: translations.file_Upload_Success
                  });
                  updateStatus(fileName, documentid, "Success", "");
                }
                if (chatId && documentid) {
                  console.log("................");
                  let favchats = [];
                  let allchatsRec = recentChats.map((chatitem) => {
                    if (chatitem.id === chatId) {
                      chatitem.document_id = [
                        ...chatitem.document_id,
                        documentid,
                      ];
                      chatitem.original_file_name = [
                        ...chatitem.original_file_name,
                        fileName,
                      ];
                    }
                    if (chatitem.isFavorite === true) {
                      favchats.push(chatitem);
                    }
                    return chatitem;
                  });

                  setrecentChats(allchatsRec);
                  dispatch(setRecentChats(allchatsRec));
                  setFilteredRecentChats(allchatsRec);
                  setFavRecentChats(favchats);
                  setFilteredFavRecentChats(favchats);
                }
              })
              .catch((error) => {
                console.log(error, fileName);
                errorList.push({
                  filename: fileName,
                  status: "Failure",
                  message: error.message,
                });
                if (
                  localStorage.getItem("chat_id") === chatId ||
                  localStorage.getItem("chatsessionid") === requestsession
                ) {
                  updateStatus(fileName, documentid, "Failed", error.message);
                }
              });
          })
          .catch((error) => {
            // hanlde erros in  predisn or upload
            console.log(error, "from presigned");
            errorList.push({
              filename: fileName,
              Status: "Failure",
              message: error.message,
            });

            if (
              localStorage.getItem("chat_id") === chatId ||
              localStorage.getItem("chatsessionid") === requestsession
            ) {
              updateStatus(fileName, "", "Failed", error.message);
            }
          });
      });

      console.log(uploadPromises);
      Promise.allSettled(uploadPromises).then((x) => {
        console.log(x);
        console.log(errorList, successList);
        if (
          localStorage.getItem("chat_id") === chatId ||
          localStorage.getItem("chatsessionid") === requestsession
        ) {
          setdocument_id(documentIdList);
          setIsFileProcessing(false);
          setisResponseDone(true);
        }
      });
    }
  };

  const pollDocumentStatus = (
    chat_id,
    documentStatusLength,
    attempts = 0,
    maxAttempts = 90,
    interval = 3000
  ) => {
    return new Promise((resolve, reject) => {
      const intervalID = setInterval(() => {
        fnGetStatusForDocument(azureUser.accessToken, chat_id)
          .then((response) => {
            const docs = response?.documents || [];
            let allCompleted = true;

            docs.forEach((doc) => {
              const {
                document_id,
                document_status,
                failure_reason,
                message_id,
              } = doc;

              if (document_status === "Success") {
                updateStatus(
                  selectedPrompt==="EL1" ? message_id: message_id - 1,
                  document_id,
                  "Success",
                  "",
                  "uploaded"
                );
              } else if (document_status === "Failed") {
                updateStatus(
                  selectedPrompt==="EL1" ? message_id: message_id - 1,
                  document_id,
                  "Failed",
                  failure_reason,
                  "Failed",
                );
                if (chatInputErrorMessage.status === false)
                  setChatInputErrorMessage({
                    status: true,
                    message: (failure_reason ?? "") + "  " + "Request Id: " + response.lastRequestConfig
                  });
              } else {
                allCompleted = false;
              }
            });

            if (allCompleted && docs.length === documentStatusLength) {
              clearInterval(intervalID);
              resolve("success");
            } else if (attempts > maxAttempts) {
              clearInterval(intervalID);
              // Only update files in 'In progress' state to Failed with timeout message
              setDocumentStatusList((prev) =>
                prev.map((item) =>
                  item.state === "In progress" || item.state === "pending"
                    ? {
                      ...item,
                      status: "Failed",
                      message: "Polling timed out. File not uploaded.",
                      state: "Failed",
                    }
                    : item
                )
              );
              setFilesToUpload((prev) =>
                prev.map((item) =>
                  item.state === "In progress" || item.state === "pending"
                    ? {
                      ...item,
                      status: "Failed",
                      message: "Polling timed out. File not uploaded.",
                      state: "Failed",
                    }
                    : item
                )
              );
              reject(new Error("Request timed out"));
            }
          })
          .catch((err) => {
            clearInterval(intervalID);
            reject(err);
          });
        attempts++;
      }, interval);
    });
  };

  const fnMutliUploadHandle_v2 = (
    files,
    chat_id,
    session,
    pollingStatusDocumentLength
  ) => {
    setisResponseDone(false);
    setIsFileProcessing(true);
    setChatInputErrorMessage({ status: false, message: "" });
    if (isHome) setHome(false);
    if (selectedPrompt.length === 0) setshowWelcomeMessage(true);

    function check_Valid_session(chat_id, session) {
      return (
        localStorage.getItem("chat_id") === chat_id ||
        localStorage.getItem("chatsessionid") === session
      );
    }

    function extract_headers(presigned_response) {
      let extracted_headers = {};
      Object.keys(presigned_response).forEach((key) =>
        key !== "urls"
          ? (extracted_headers[key] = presigned_response[key])
          : null
      );
      return extracted_headers;
    }

    function error_Handling_state_upload(chat_id, session, error) {
      if (check_Valid_session(chat_id, session)) {
        setIsFileProcessing(false);
        setisResponseDone(true);
        setFilesToUpload((prev) =>
          prev.map((item) =>
            item.state === "pending"
              ? {
                ...item,
                status: "Failed",
                message:
                  error?.message ?? "Error in presigned url generation",
                state: "Failed",
              }
              : item
          )
        );
        setDocumentStatusList((prev) =>
          prev.map((item) =>
            item.state === "pending"
              ? {
                ...item,
                status: "Failed",
                message:
                  error?.message ?? "Error in presigned url generation",
                state: "Failed",
              }
              : item
          )
        );
      }
    }

    const updatedoc_id = (key, doc_id, status, message) => {
      setDocumentStatusList((prev) =>
        prev.map((item) =>
          item.key === key ? { ...item, status, message, doc_id } : item
        )
      );
      setFilesToUpload((prev) =>
        prev.map((item) =>
          item.key === key ? { ...item, status, message, doc_id } : item
        )
      );
    };

    if (check_Valid_session(chat_id, session)) {
      let Filenames = files.map((file) => file.file.name);
      fnGetPresingedURL(
        azureUser.accessToken,
        Filenames,
        chat_id,
        folder_id,
        azureUser.selected_model.model_id,
        false,
        null,
        selectedPrompt === "EL1"
      )
        .then((response) => {
          if (response?.chat_id) {
            setchatID(response.chat_id);
            localStorage.setItem("chat_id", response.chat_id);
          }
          const presignedUrls = response.urls;
          if (presignedUrls.length === 0) {
            throw new Error("Error in Generating the presigned urls" + " " + "Request Id: " + response.lastRequestConfig);
          }
          const put_headers_json = extract_headers(response);

          const uploadPromises = files.map((file, index) => {
            let fileName = file.file.name;
            let fileData = file.file;
            const { presigned_url, document_id, file_name_full } =
              presignedUrls[index];

            updatedoc_id(file.key, document_id, "In progress", "");

            let filenameHeader = fileName;
            if (file_name_full !== fileName) {
              filenameHeader = file_name_full;
              fileData = new File([file.file], file_name_full, {
                type: file.file.type,
              });
            }
            const headers_for_put = new Headers();
            headers_for_put.append("Content-Type", file.file.type);
            headers_for_put.append("x-amz-meta-filename", filenameHeader);
            headers_for_put.append("x-amz-meta-document_id", document_id);
            if(selectedPrompt === "EL1")
                headers_for_put.append("x-amz-meta-is_elgpt_upload", "True");
            Object.keys(put_headers_json || {}).forEach((key) => {
              headers_for_put.append(
                `x-amz-meta-${key}`,
                put_headers_json[key]
              );
            });

            return fetch(presigned_url, {
              method: "PUT",
              headers: headers_for_put,
              body: fileData,
            })
              .then(async (uploadResponse) => {
                const statuscode = uploadResponse.status;
                if (uploadResponse.ok) {
                  // Success
                } else {
                  let errorMsg = "File upload failed. Please try again.";
                  try {
                    const errorBody = await uploadResponse.text();
                    if (errorBody) errorMsg = errorBody;
                  } catch { }
                  if (statuscode === 400) {
                    throw new Error(errorMsg);
                  } else if (statuscode === 403)
                    throw new Error(
                      "Unable to verify token. Please try again."
                    );
                  else throw new Error(errorMsg);
                }
              })
              .catch((error) => {
                updatedoc_id(fileName, document_id, "Failed", error.message);
                setChatInputErrorMessage({
                  status: true,
                  message: error.message,
                });
              });
          });

          Promise.allSettled(uploadPromises)
            .then((result) => {
              const allsuccess = result.every((r) => r.status === "fulfilled");
              return allsuccess;
            })
            .then((allsuccess) => {
              if (allsuccess) {
                pollDocumentStatus(
                  response?.chat_id,
                  pollingStatusDocumentLength
                )
                  .then(() => {
                    setIsFileProcessing(false);
                    setisResponseDone(true);
                  })
                  .catch((err) => {
                    setChatInputErrorMessage({
                      status: true,
                      message:
                        err?.message ??
                        "Unidentified Error occurred, please try again",
                    });

                    setFilesToUpload((prev) =>
                      prev.map((item) =>
                        item.state !== "uploaded"
                          ? {
                            ...item,
                            status: "Failed",
                            message: "Polling error. File not uploaded.",
                            state: "Failed",
                          }
                          : item
                      )
                    );
                    setDocumentStatusList((prev) =>
                      prev.map((item) =>
                        item.state !== "uploaded"
                          ? {
                            ...item,
                            status: "Failed",
                            message: "Polling error. File not uploaded.",
                            state: "Failed",
                          }
                          : item
                      )
                    );
                    setIsFileProcessing(false);
                    setisResponseDone(true);
                  });
              } else {
                setFilesToUpload((prev) =>
                  prev.map((item) =>
                    item.state !== "uploaded"
                      ? {
                        ...item,
                        status: "Failed",
                        message: "One or more files failed to upload.",
                        state: "Failed",
                      }
                      : item
                  )
                );
                setDocumentStatusList((prev) =>
                  prev.map((item) =>
                    item.state !== "uploaded"
                      ? {
                        ...item,
                        status: "Failed",
                        message: "One or more files failed to upload.",
                        state: "Failed",
                      }
                      : item
                  )
                );
                setChatInputErrorMessage({
                  status: true,
                  message: "Some files failed to upload.",
                });
                setIsFileProcessing(false);
                setisResponseDone(true);
              }
            })
            .catch((err) => {
              setIsFileProcessing(false);
              setisResponseDone(true);
              error_Handling_state_upload(chat_id, session, err);
            });
        })
        .catch((error) => {
          error_Handling_state_upload(chat_id, session, error);
          setChatInputErrorMessage({
            status: true,
            message: translations?.Presigned_url_error,
          });
          setIsFileProcessing(false);
          setisResponseDone(true);
        })
        .finally(() => {
          setFileUploadComponentErros({ type: "", message: "", status: false });
        });
    }
  };

  const onFilesValidated = (files, pollingStatusDocumentLength) => {
    fnMutliUploadHandle_v2(
      files,
      chatID,
      localStorage.getItem("chatsessionid"),
      pollingStatusDocumentLength
    );
  };

  const pollDocumentJobStatus = (
    requestsession,
    chatID,
    job_id,
    attempts = 0,
    maxAttempts = 120
  ) => {
    return new Promise((resolve, reject) => {
      const intervalID = setInterval(() => {
        fnGetSummaryandTranslateDocumentJobStatus(
          azureUser.accessToken,
          chatID,
          job_id
        )
          .then((response) => response.json())
          .then((data) => {
            console.log(data);
            if (data?.job_status === "ready") {
              console.log("entered if for job status ready");
              clearInterval(intervalID);
              resolve("success");
            } else if (data?.job_status?.toLowerCase() === "failed") {
              clearInterval(intervalID);
              reject(new Error("Failed"));
            } else if (
              localStorage.getItem("chatsessionid") !== requestsession
            ) {
              clearInterval(intervalID);
              reject(new Error("aborted"));
            } else if (attempts >= maxAttempts) {
              clearInterval(intervalID);
              reject(new Error("Request timed out"));
            }
          })
          .catch((err) => {
            clearInterval(intervalID);
            reject(err);
          });
        if (maxAttempts) attempts++;
      }, 5000);
    });
  };

  const fnEditTitle = (e, chatid) => {
    console.log(e);
    console.log(chatid);
    setnewTitle("");
    setisChangeClicked(true);
    seteditChatTitleID(chatid);
    setTimeout(
      () => titleRef.current.shadowRoot.querySelector("textarea").focus(),
      500
    );
    console.log("edit title", e.target);
  };

  const fnChangeTitle = () => {
    console.log("change title", editChatTitleID, newTitle);
    fnChangeChatTitle(azureUser.accessToken, editChatTitleID, newTitle)
      .then((response) => {
        console.log("response1", response);
        setdispAlertMessage({
          status: true,
          code: "success",
          message: translations?.Title_changed,
        });
        setisChangeClicked(false);
        setrecentChats([]);
        fnLoadRecentChats(azureUser.accessToken);
        setnewTitle("");
      })
      .catch((error) => {
        console.log("Error in fnChangeTitle", error)
        setdispAlertMessage({
          status: true,
          code: "error",
          message: "Unable to change title. Please try again.",
        });
      });
  };

  const fnMoveChat = (chat_id, target_folder_id) => {
    fnChatUpdateFolder(azureUser.accessToken, chat_id, target_folder_id)
      .then((response) => {
        if (response) {
          console.log("response1", response);
          setdispAlertMessage({
            status: true,
            code: "success",
            message: `${translations?.chat_moved}`,
          });
          setisChangeClicked(false);
          setrecentChats([]);
          fnLoadRecentChats(azureUser.accessToken);
        } else {
          throw new Error(translations?.unable_to_move);
        }
      })
      .catch((error) => {
        console.log(error);
        setdispAlertMessage({
          status: true,
          code: "error",
          message: error.message,
        });
      });
  };
  const fnDeleteChat = (e, chatid) => {
    console.log("delete title", chatid);
    setisDeleteClicked(true);
    setdeleteChatID(chatid);
  };

  const fnRemoveChat = () => {
    console.log("delete title", deleteChatID);
    fnDeleteChatID(azureUser.accessToken, deleteChatID)
      .then((response) => {
        console.log("response1", response);
        setdispAlertMessage({
          status: true,
          code: "success",
          message: translations?.Delete_chat_notification,
        });
        setisDeleteClicked(false);
        fnLoadRecentChats(azureUser.accessToken);
        console.log(
          "localStorage.getItem() === deleteChatID",
          localStorage.getItem("chat_id"),
          deleteChatID,
          localStorage.getItem("chat_id") === deleteChatID
        );

        if (localStorage.getItem("chat_id") === deleteChatID) {
          setrecentChats([]);
          fnLoadRecentChats(azureUser.accessToken);
          setdeleteChatID("");
          localStorage.setItem("chat_id", "");
          setchatID();
          setresponseValues([]);
          setFileList([]);
          setdocument_id([]);
          setBlocked_chat_context({
            is_blocked: false,
            blocked_message: "",
          });
          inputQueryRef.current.value = "";
          //setMessage("");
          setTimeout(() => {
            const textarea = document.getElementById("inputQuery");
            textarea.style.height = "60px";
          }, 500);
          setDocumentStatusList([]);
          setFileUploadComponentErros({ type: "", message: "", status: false });
          setNewChatLoading(true);
          setshowWelcomeMessage(false);
          setTimeout(() => {
            setHome(true);
            setNewChatLoading(false);
          }, 2000);
        }
      })
      .catch((error) => {
        console.log("Error in fnRemoveChat", error)
        setdispAlertMessage({
          status: true,
          code: "error",
          message: "Unable to delete. Please try again.",
        });
      });
  };

  const fnShowDetails = (event, searchValue) => {
    console.log("first");
    console.log(event, searchValue);
    if (searchValue === translations?.Frequently_asked) {
      setviewStaticData({
        isClicked: true,
        clickedItem: searchValue,
        Data: CapabilitiesArray,
      });
    } else if (searchValue === translations?.Prompt_Guide) {
      setviewStaticData({
        isClicked: true,
        clickedItem: searchValue,
        Data: [],
      });
    }
  };

  const fnAddFav = (event) => {
    let chatID = event.target.getAttribute("chat_id");
    setFavMessageLoad(true);
    let isF = isFav;
    setisFav(isF ? false : true);

    fnChangeChatFav(azureUser.accessToken, chatID, isF ? false : true)
      .then((response) => {
        console.log("response1", response);
        // setdispAlertMessage({
        //   status: true,
        //   code: "success",
        //   message: "Favourite updated.",
        // });
        setisChangeClicked(false);
        setrecentChats([]);
        fnLoadRecentChats(azureUser.accessToken);
        setnewTitle("");
      })
      .catch((error) => {
        console.log("Error in fnAddFav", error)
        setdispAlertMessage({
          status: true,
          code: "error",
          message: "Unable to udpate favourite. Please try again.",
        });
      });
  };

  const fnAcceptedOptOut = (e, opt_in, isnoticeRead) => {
    fnPOSTUsersOptStatus(azureUser.accessToken, opt_in, isnoticeRead)
      .then((res) => {
        console.log(res);

        if (res?.opt_in) {
          setdispAlertMessage({
            status: true,
            code: "error",
            message: "Unable to opt out. please try again",
          });
        } else {
          setoptOutWarning(false);
          let useravatarContainer = document
            .querySelector("sfc-shell-app-bar")
            .shadowRoot.querySelector(".logged-in-user")
            .querySelector("sdf-floating-pane")
            .querySelector(".avatar-menu")
            .querySelector("sdf-switch");
          useravatarContainer.setAttribute("checked", "false");
          useravatarContainer.setAttribute("disabled", "true");
          useravatarContainer.setAttribute("label", "Opt-out");
          setdispAlertMessage({
            status: true,
            code: "success",
            message: "You have been opted out.",
          });

          setazureUser((prev) => {
            return {
              ...prev,
              isPreviledgedUser: true,
            };
          });

          console.log("home clicked");
          setchatID();
          localStorage.setItem("chat_id", "");
          setresponseValues([]);
          inputQueryRef.current.value = "";
          //setMessage("");
          setTimeout(() => {
            const textarea = document.getElementById("inputQuery");
            textarea.style.height = "60px";
          }, 500);
          setHome(false);
          setdocument_id([]);
          setBlocked_chat_context({
            is_blocked: false,
            blocked_message: "",
          });
          setFileList([]);
          setDocumentStatusList([]);
          setFilesToUpload([]);
          setFileUploadComponentErros({ type: "", message: "", status: false });
          setNewChatLoading(true);
          fnLoadRecentChats(azureUser.accessToken);
          setshowWelcomeMessage(false);
          setTimeout(() => {
            setHome(true);
            setNewChatLoading(false);
          }, 2000);
        }
      })
      .catch((error) => {
        console.log(error);
        let useravatarContainer = document
          .querySelector("sfc-shell-app-bar")
          .shadowRoot.querySelector(".logged-in-user")
          .querySelector("sdf-floating-pane")
          .querySelector(".avatar-menu")
          .querySelector("sdf-switch");
        useravatarContainer.setAttribute("checked", "true");

        setdispAlertMessage({
          status: true,
          code: "error",
          message: "Unable to opt out. please try again",
        });
      });
  };

  const fnuserNoticeAcknowledgment = (event, opt_in, isnoticeRead) => {
    fnPOSTUsersOptStatus(azureUser.accessToken, opt_in, isnoticeRead)
      .then((res) => {
        console.log("print res", res)
        setazureUser({ ...azureUser, isnoticeRead: true });
      })
      .error((err) => console.log(err));
  };

  const chatNavigationHandler = async (e, context) => {
    console.log("clicked chat nav", context);
    console.log(recentChats);
    const notificationchat = await recentChats.filter(
      (item) => item.id === context.chat_id
    );

    if (notificationchat.length > 0) {
      const blocked_chat_context = {
        is_blocked: notificationchat?.[0]?.block_chat ?? false,
        blocked_message: notificationchat?.[0]?.block_message ?? "",
      };
      fnLoadChat(
        e,
        context?.chat_id,
        notificationchat?.[0]?.isFavorite,
        notificationchat?.[0]?.document_id,
        notificationchat?.[0]?.original_file_name,
        notificationchat?.[0]?.model_id,
        notificationchat?.[0]?.llm_name,
        blocked_chat_context
      );
    } else {
      const model_id = context?.message?.llmResponse?.model_id;
      console.log(azureUser);
      const model = azureUser.models_accessible.filter(
        (model) => model.model_id === model_id
      );

      const blocked_chat_context = {
        is_blocked: false,
        blocked_message: "",
      };
      console.log(model);
      fnLoadChat(
        e,
        context?.chat_id,
        false,
        [],
        [],
        model_id,
        model?.[0]?.llm_name,
        blocked_chat_context
      );
    }

    setMessageReadyNotificationContext({
      is_message_ready: false,
      // chat_title: "",
      // chat_id: ""
    });
  };
  const fnAddFolderModal = () => {
    setFolderTitle("");
    SetIsAddFolderClicked(true);
    setFolder_id();
    setTimeout(
      () => foldertitleRef.current.shadowRoot.querySelector("textarea").focus(),
      500
    );
  };

  const fnEditFolderModal = (e, f_id) => {
    console.log("model called............", e, f_id);
    setFolder_id(f_id);
    setFolderTitle("");
    setIsupdateFolderClicked(true);
    setTimeout(
      () =>
        editfoldertitleRef.current.shadowRoot.querySelector("textarea").focus(),
      500
    );
  };

  function fnHandleAddFolder() {
    fnAddFolder(azureUser.accessToken, folderTitle)
      .then((response) => {
        console.log("response1", response);
        if (response) {
          setdispAlertMessage({
            status: true,
            code: "success",
            message: `${translations?.Folder_added}`,
          });
          SetIsAddFolderClicked(false);
          setFolder_id(response?.folder_details?.folder_id);
          setrecentChats([]);
          fnLoadRecentChats(azureUser.accessToken);
          setFolderTitle("");
          fnNewChat();
        } else {
          throw new Error("Unable to add a folder. Please try again.");
        }
      })
      .catch((error) => {
        console.log(error);
        setdispAlertMessage({
          status: true,
          code: "error",
          message: error.message,
        });
      });
  }

  function fnHandleUpdateFolder() {
    fnEditFolderTitle(azureUser.accessToken, folder_id, folderTitle)
      .then((response) => {
        console.log("response1", response);
        if (response) {
          setdispAlertMessage({
            status: true,
            code: "success",
            message: translations?.folder_title_changes,
          });
          setIsupdateFolderClicked(false);
          //setFolder_id(response?.folder_details?.folder_id)
          fnLoadRecentChats(azureUser.accessToken);
          setFolderTitle("");
          //  fnNewChat()
        } else {
          throw new Error("Unable to edit the folder. Please try again.");
        }
      })
      .catch((error) => {
        setdispAlertMessage({
          status: true,
          code: "error",
          message: error.message,
        });
      });
  }

  function fnHandleDeleteFolder() {
    fnDeletFolder(azureUser.accessToken, deleteFolderId)
      .then((response) => {
        console.log("response1", response);
        setdispAlertMessage({
          status: true,
          code: "success",
          message: translations?.folder_deleted,
        });
        setIsDeleteFolderClicked(false);
        fnLoadRecentChats(azureUser.accessToken);
      })
      .catch((error) => {
        console.log("Error in fnHandleDeleteFolder", error)
        setdispAlertMessage({
          status: true,
          code: "error",
          message: "Unable to delete. Please try again.",
        });
      });
  }

  //voice code below

  const voiceStore = useSelector((state) => state.voice);
  const [microphoneAvailable, setmicrophoneAvailable] = useState(true);
  async function checkMicrophoneAvailable() {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      stream.getTracks().forEach((track) => track.stop());
      return true;
    } catch (err) {
      return false;
    }
  }

  const resetInlineVoiceChat = () => {
    setListeningWaveForm(false);
    setTimeout(() => {
      const textarea = document.getElementById("inputQuery");
      if (textarea) textarea.style.height = "58px";
    }, 500);
    set_voice_mode("voice_inline");
    // fnHandleSubmit(localStorage.getItem("chatsessionid"));
    stopListening();
    dispatch && dispatch({ type: "voice/setStopListening" });

    resetTranscript();
    dispatch(setTranscript(""));
    // setInputRefVal("");
  };

  const {
    transcript,
    resetTranscript,
    listening,
    browserSupportsSpeechRecognition,
  } = useSpeechRecognition({});

  useEffect(() => {
    if (inputQueryRef && inputQueryRef.current)
      inputQueryRef.current.value = transcript;
  }, [transcript]);
  //   Copied commented by Umesh for future release
  useEffect(() => {
    const markdownContainers = document.querySelectorAll(".snippet-container .react-markdown");

    markdownContainers.forEach((container) => {
      const preBlocks = container.querySelectorAll("pre");

      preBlocks.forEach((pre) => {
        if (pre.querySelector(".copy-btn")) return;

        const copyIconSvgString = `
        <svg
          xmlns="http://www.w3.org/2000/svg"
          viewBox="0 0 24 24"
          fill="transparent"
          stroke="white"
          stroke-width="2"
          stroke-linecap="round"
          stroke-linejoin="round"
          class="w-4 h-4"
        >
          <path d="M0 0h24v24H0z" fill="transparent" stroke="none" />
          <path d="M7 7m0 2.667a2.667 2.667 0 0 1 2.667 -2.667h8.666a2.667 2.667 0 0 1 2.667 2.667v8.666a2.667 2.667 0 0 1 -2.667 2.667h-8.666a2.667 2.667 0 0 1 -2.667 -2.667z" />
          <path d="M4.012 16.737a2.005 2.005 0 0 1 -1.012 -1.737v-10c0 -1.1 .9 -2 2 -2h10c.75 0 1.158 .385 1.5 1" />
        </svg>
      `;
        const button = document.createElement("button");
        button.className =
          "copy-btn absolute top-2 right-2 text-black p-0.5 rounded-md flex items-center justify-center";
        button.innerHTML = copyIconSvgString;

        button.onclick = () => {
          const codeText = pre.querySelector("code")?.innerText || "";
          navigator.clipboard.writeText(codeText).then(() => {
            button.innerText = "Copied!";
            button.classList.add("bg-green-100", "text-green-600");
            setTimeout(() => {
              button.innerHTML = copyIconSvgString;
              button.classList.remove("bg-green-100", "text-green-600");
            }, 2500);
          });
        };

        pre.style.position = "relative";
        pre.appendChild(button);
      });
    });
  }, [responseValues]);

  const startListening = async () => {
    console.log("Started Listening in ", voiceStore.language);
    SpeechRecognition.startListening({
      continuous: true,
      language: voiceStore.language,
      // language: 'hi-IN'

      onEnd: () => console.log("speech ended"),
    });
  };

  const stopListening = () => {
    console.log("stop listening library");
    SpeechRecognition.stopListening();
  };

  const [isMuted, setIsMuted] = useState(false);
  const [voiceDisplayText, setvoiceDisplayText] = useState("Listening");
  const [disableMute, setDisableMute] = useState(false);

  function setInputRefVal(value) {
    inputQueryRef.current.value = value;
  }

  let utterance = null;

  const [pollyAudio] = useState();

  const controls = useAnimation();

  const fnDeleteFolderModalOpener = (e, f_id) => {
    setDeletedFolderId(f_id);
    setIsDeleteFolderClicked(true);
  };

  function Handle_Remove_SuccessFiles(file, requestSession) {
    setisResponseDone(false);
    setIsFileProcessing(true);
    setChatInputErrorMessage({
      status: false,
      message: "",
    });
    console.log(file, "uploaded file removal");
    fnDeleteSuccessFile(azureUser.accessToken, chatID, file?.[0]?.doc_id)
      .then((response) => {
        const status_code = response.status;
        console.log(status_code);
        if (status_code !== 200) {
          if (localStorage.setItem("chatsessionid") === requestSession) {
            throw new Error("unable to delete the file");
          }
        }
        return response.json();
      })
      .then((data) => {
        console.log(data, file);
        Handle_Remove_inlineFile(file?.[0]);
      })
      .catch((e) => {
        console.log(e);
        setdisplayErrorMessage({
          status: true,
          message: e.message,
        });
      })
      .finally(() => {
        setisResponseDone(true);
        setIsFileProcessing(false);
      });
  }

  function Handle_Remove_inlineFile(file) {
    if (
      file.status === "Success" ||
      file.status === "Failed" ||
      file.state === "pending"
    ) {
      const filtered_filesToUPload = filesToUpload.filter(
        (f) => f.key !== file.key
      );
      const filtered_documentStatusList = documentStatusList.filter(
        (f) => f.key !== file.key
      );

      setFilesToUpload(filtered_filesToUPload);
      setDocumentStatusList(filtered_documentStatusList);

      const current_selected_files = filtered_filesToUPload.filter(
        (f) => f.status !== "Success" && f.state !== "Failed"
      );
      const isValid =
        current_selected_files.length > 0
          ? current_selected_files.every((f) => f.valid)
          : true;

      const excelType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

      const csvtype = "text/csv";

      const hasExcelAlready = filtered_documentStatusList.some(
        (f) => f.file?.type === excelType || f.file?.type === csvtype
      );
      const hasNonExcelAlready = filtered_documentStatusList.some(
        (f) => f.file?.type !== excelType && f.file?.type !== csvtype
      );

      const errors = [];
      if (selectedPrompt !="EL1" &&  (hasExcelAlready && hasNonExcelAlready) ){
        errors.push({
          type: "filetype",
          message: "Upload_limit_message",
        });
      }
      if (filtered_documentStatusList.length > MAX_FIlES) {
        errors.push({
          type: "filecount",
          message: `Only ${Math.abs(
            filtered_documentStatusList.length - MAX_FIlES
          )} files are allowed`,
        });
      }

      if (isValid && errors.length === 0) {
        if (current_selected_files.length > 0) {
          onFilesValidated(
            current_selected_files,
            filtered_documentStatusList.length
          );
        }
        setFileUploadComponentErros({
          type: "",
          message: "",
          status: false,
        });
      } else {
        if (errors.length > 0) {
          const fileTypeError = errors.find((e) => e.type === "filetype");
          if (fileTypeError) {
            setFileUploadComponentErros({
              type: "filetype",
              message: fileTypeError.message,
              status: true,
              hasprevalidationError: true,
            });
          } else {
            setFileUploadComponentErros({
              type: "filecount",
              message: `File count is exceeded, please remove the files accordingly!`,
              status: true,
              hasprevalidationError: true,
            });
          }
        } else {
          const sizeLimitError = filtered_documentStatusList.find((e) =>
            e.message.includes("exceeded size limit")
          );
          if (sizeLimitError) {
            setFileUploadComponentErros({
              type: "general",
              message: sizeLimitError.message,
              status: true,
              hasprevalidationError: true,
            });
          } else {
            setFileUploadComponentErros({
              type: "general",
              message:
                "File type is not allowed , please remove the files accordlingly",
              status: true,
              hasprevalidationError: true,
            });
          }
        }
      }
    } else {
      setFileUploadComponentErros({
        type: "delete",
        message: "Cannot delete files that are currently uploading.",
        status: true,
      });
    }
  }

  function fnactionOpenUrl(url) {
    console.log(url);

    const link = document.createElement("a");
    link.href = url;
    link.target = "_blank";
    link.click();
  }

  const OpenHelpModal = () => {
    localStorage.setItem("faq_chatId", "")
    setHelpVis(!helpVis)
  }

  const [gptData, setGptData] = useState([]);
  const [myGptDataLoader, setMyGptDataLoader] = useState(false);

  const handleSetGPTdata = () => {
    setMyGptDataLoader(translations?.Fetching_gpt_data);

    fnGetMyGpts(azureUser?.accessToken)
      .then((response) => {
        setGptData(response);
      })
      .catch((error) => {
        console.error("Error fetching GPTs:", error);
        setGptData([]);
        setdispAlertMessage({
          status: true,
          code: "error",
          message: "Something went wrong while fetching gpt data..."
        });
      })
      .finally(() => {
        setMyGptDataLoader();
      });
  }

  const [createNoteVis,setcreateNoteVis] = useState(false)

  return (
    <>
      <div className="flex relative">
        <SideBar
          {...{
            chatHistoryLoading,
            openSideNav,
            setOpenSideNav,
            openGPTScreen,
            setOpenGPTScreen,
            newOpenGPTScreen,
            setNewOpenGPTScreen,
            filteredRecentChats,
            filteredFavRecentChats,
            searchRef,
            allchatsRec,
            fnAddFolderModal,
            allFolders,
            setChatsSectionNavState,
            fnMoveChat,
            fnNewChat,
            fnLoadChat,
            setOpenSearch,
            openSearch,
            fnEditFolderModal,
            fnDeleteChat,
            fnEditTitle,
            fnDeleteFolderModalOpener,
            folder_id,
            setFolder_id,
            Model_state_rest, 
            darkModeVars,
            handleSetGPTdata,
            openAIGIFlow, 
            setOpenAIGIFlow,
            setOpenBPQ,
            setOpenFIJ          
          }}
        />

 <CreateNotification

          visible={createNoteVis}
          onClose={() => setcreateNoteVis(false)}
          onSubmit={(notificationData) => {
            fnCreateNotification(azureUser?.accessToken,notificationData)
            .then((response)=>{
              setdispAlertMessage({
                status: true,
                code: "success",
                message: translations?.Notification_success,
              });
            })
            .catch((error)=>{
              setdispAlertMessage({
                status: true,
                code: "error",
                message: error?.response?.data?.message || translations?.Backend_error,
              });
            })
            
          }}
          {...{setdispAlertMessage,createNoteRef}}
        />
        

        {sessionGroupsModal.visible && (
          <Webex_Groups_Search
            searchRef_webex_groups={searchRef_webex_groups}
            setSessionGroupsModal={setSessionGroupsModal}
            sessionGroupsModal={sessionGroupsModal}
            selectedPrompt={selectedPrompt}
            fnHandleSubmit={fnHandleSubmit}
          ></Webex_Groups_Search>
        )}
        <SdfFocusPane
          size="small"
          heading={translations.Opt_out_confirmation}
          visible={optOutWarning}
          closeable={true}
          onSdfAccept={(e) =>
            fnAcceptedOptOut(e, false, azureUser.isnoticeRead)
          }
          acceptLabel="Yes"
          dismissLabel="No"
          onSdfDismiss={() => {
            setoptOutWarning(false);

            let useravatarContainer = document
              .querySelector("sfc-shell-app-bar")
              .shadowRoot.querySelector(".logged-in-user")
              .querySelector("sdf-floating-pane")
              .querySelector(".avatar-menu")
              .querySelector("sdf-switch");
            useravatarContainer.setAttribute("checked", "true");
          }}
        >
          <div className="p-5">{process.env.REACT_APP_OPTOUT_WARNING}</div>
        </SdfFocusPane>

        <SdfFocusPane
          size="medium"
          heading={translations.Select_prompts}
          hideAcceptButton={true}
          hideDismissButton={true}
          visible={isPromptClicked}
          closeable={true}
          onSdfDismiss={() => {
            localStorage.setItem("LatestToRender", "");
            setisPromptClicked(false);
          }}
        >
          <div className="col-start-2 col-end-12 grid grid-cols-12 gap-3 p-4">
            {documentStatusList.length === 0 ? (
              <>
                {promptsData.map((valP) => {
                  return (
                    <>
                      {valP.prompt_mode === "general" ||
                        valP.prompt_mode === "user_defined" ||
                        valP.prompt_mode === "Webex_Agent" ? (
                        <div
                          onClick={(e) =>
                            fnClickPrompt(
                              e,
                              valP.prompt_name,
                              valP.prompt_id,
                              valP.prompt_mode,
                              localStorage.getItem("chatsessionid")
                            )
                          }
                          title={valP.prompt_name}
                          prompt_id={valP.prompt_id}
                          className="cursor-pointer flex flex-col items-start enhancementsCards  col-span-6"
                        >
                          <div
                            className="w-64 gap-2"
                            title={valP.prompt_name}
                            prompt_id={valP.prompt_id}
                          >
                            <div
                              className="flex justify-center"
                              title={valP.prompt_name}
                              prompt_id={valP.prompt_id}
                            >
                              <SdfSpotIllustration
                                illustration-name={
                                  illustrationNames[valP.prompt_name]
                                }
                                title={valP.prompt_name}
                                prompt_id={valP.prompt_id}
                              ></SdfSpotIllustration>
                            </div>
                            <div
                              className="flex justify-center"
                              title={valP.prompt_name}
                              prompt_id={valP.prompt_id}
                            >
                              <b>{valP.prompt_name}</b>
                            </div>
                            <br />
                            <div
                              className="flex justify-center"
                              title={valP.prompt_name}
                              prompt_id={valP.prompt_id}
                            >
                              <p style={{ fontSize: "10px" }}>
                                {valP.prompt_tagline}
                              </p>
                            </div>
                          </div>
                        </div>
                      ) : (
                        <></>
                      )}
                    </>
                  );
                })}
              </>
            ) : (
              <>
                {promptsData.map((valP) => {
                  return (
                    <>
                      {valP.prompt_mode === "document_upload" ? (
                        <div
                          onClick={
                            valP.prompt_name === "Summarize Documents"
                              ? () =>
                                fnHandleSubmit(
                                  localStorage.getItem("chatsessionid"),
                                  "Summarize Documents",
                                  valP.prompt_user_utterance
                                )
                              : () => {
                                console.log("input query translation");
                                inputQueryRef.current.value =
                                  valP.prompt_user_utterance;
                                //setMessage(valP.prompt_user_utterance);
                                setisPromptClicked(false);
                              }
                          }
                          title={valP.prompt_name}
                          prompt_id={valP.prompt_id}
                          className="cursor-pointer flex flex-col items-start enhancementsCards  col-span-6"
                        >
                          <div
                            className="w-64 gap-2"
                            title={valP.prompt_name}
                            prompt_id={valP.prompt_id}
                          >
                            <div
                              className="flex justify-center"
                              title={valP.prompt_name}
                              prompt_id={valP.prompt_id}
                            >
                              <SdfSpotIllustration
                                illustration-name={
                                  illustrationNames[valP.prompt_name]
                                }
                                title={valP.prompt_name}
                                prompt_id={valP.prompt_id}
                              ></SdfSpotIllustration>
                            </div>
                            <div
                              className="flex justify-center"
                              title={valP.prompt_name}
                              prompt_id={valP.prompt_id}
                            >
                              <b>{valP.prompt_name}</b>
                            </div>
                            <br />
                            <div
                              className="flex justify-center"
                              title={valP.prompt_name}
                              prompt_id={valP.prompt_id}
                            >
                              <p style={{ fontSize: "10px" }}>
                                {valP.prompt_tagline}
                              </p>
                            </div>
                          </div>
                        </div>
                      ) : (
                        <></>
                      )}
                    </>
                  );
                })}
              </>
            )}

          </div>
        </SdfFocusPane>

        <SdfFocusPane
          size="large"
          heading={viewStaticData.clickedItem}
          hideAcceptButton={true}
          hideDismissButton={true}
          visible={viewStaticData.isClicked}
          closeable={true}
          onSdfDismiss={() => {
            setviewStaticData({
              isClicked: false,
              clickedItem: "",
              Data: [],
            });
            setActiveTab(translations?.Basic);
            setIsPromptOpen(false);
          }}
        >
          <>
            {viewStaticData.clickedItem === translations?.Prompt_Guide ? (
              <div className=" p-2">
                <SdfTabGroup
                  className="p-2"
                  onSdfChange={(e) => {
                    setActiveTab(e.detail.value);
                    setIsPromptOpen(false);
                  }}
                >
                  <SdfTab active={activeTab === "Basic"} value={"Basic"}>
                    {translations?.Basic}
                  </SdfTab>
                  <SdfTab active={activeTab === "Advanced"} value={"Advanced"}>
                    {translations?.Advanced}
                  </SdfTab>
                </SdfTabGroup>
                {activeTab === "Basic" ? (
                  <>
                    <div className="flex justify-end pt-2 pb-2">
                      <SdfButton
                        emphasis="tertiary"
                        onClick={() => {
                          setIsPromptOpen(!isPromptOpen);
                        }}
                      >
                        {isPromptOpen ? translations?.Collapse_All : translations?.Expand_All}
                      </SdfButton>
                    </div>
                    <SdfAccordion className="h-full" allowMultiple>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Understand_header}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.Understand_pera1}
                          </li>
                          <li>{translations?.Understand_pera2}</li>
                          <li className="pl-6">
                            {translations?.Understand_pera3}: "{translations?.Understand_pera4}"
                          </li>
                          <li className="pl-6">
                            {translations?.Understand_pera5}: "{translations?.Understand_pera6}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Specific_header}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.Clearly_pera1}
                          </li>
                          <li>{translations?.Understand_pera2}</li>
                          <li className="pl-6">
                            {translations?.Clearly_pera3}: "{translations?.Clearly_pera4}"
                          </li>
                          <li className="pl-6">
                            {translations?.Clearly_pera5}: "{translations?.Clearly_pera2}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Constraints_header}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.Constraints_pera1}
                          </li>
                          <li>{translations?.Understand_pera2}</li>
                          <li className="pl-6">
                            "{translations?.Constraints_pera2}"
                          </li>
                          <li className="pl-6">
                            "{translations?.Constraints_pera3}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Context_header}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.Context_pera1}
                          </li>
                          <li>{translations?.Understand_pera2}</li>
                          <li className="pl-6">
                            "{translations?.Context_pera2}"
                          </li>
                          <li className="pl-6">
                            "{translations?.Context_pera3}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Experiment_header}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.Experiment_pera1}
                          </li>
                          <li>
                            {translations?.Experiment_pera2}
                          </li>
                          <li>{translations?.Understand_pera2}</li>
                          <li className="pl-6">
                            {translations?.Experiment_pera3}: "{translations?.Experiment_pera4}"
                          </li>
                          <li className="pl-6">
                            {translations?.Experiment_pera5}: "{translations?.Experiment_pera6}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Guide_header}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.Guide_pera1}
                          </li>
                          <li>{translations?.Understand_pera2}</li>
                          <li className="pl-6">
                            "{translations?.Guide_pera2}"
                          </li>
                          <li className="pl-6">
                            "{translations?.Guide_pera3}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Iterative_header}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.Iterative_pera1}
                          </li>
                          <li>{translations?.Understand_pera2}</li>
                          <li className="pl-6">
                            {translations?.Iterative_pera2}: "{translations?.Iterative_pera3}"
                          </li>
                          <li className="pl-6">
                            {translations?.Iterative_pera4}: "{translations?.Iterative_pera5}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Prompts_header}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.Specific_pera1} “{translations?.Specific_pera2}” {translations?.Specific_pera3}
                          </li>
                          <li>{translations?.Understand_pera2}</li>
                          <li className="pl-6">
                            "{translations?.Specific_pera4}"
                          </li>
                          <li className="pl-6">
                            "{translations?.Specific_pera5}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Leverage_header}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.Leverage_pera1}
                          </li>
                          <li>{translations?.Understand_pera2}</li>
                          <li className="pl-6">
                            "{translations?.Leverage_pera2}"
                          </li>
                          <li className="pl-6">
                            "{translations?.Leverage_pera3}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Refine_header}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.Refine_pera1}
                          </li>
                          <li>
                            {translations?.Refine_pera2}
                          </li>
                        </div>
                      </SdfExpandableBox>
                    </SdfAccordion>
                  </>
                ) : (
                  <>
                    <div className="flex justify-end pt-2 pb-2">
                      <SdfButton
                        emphasis="tertiary"
                        onClick={() => {
                          setIsPromptOpen(!isPromptOpen);
                        }}
                      >
                        {isPromptOpen ? translations?.Collapse_All : translations?.Expand_All}
                      </SdfButton>
                    </div>
                    <SdfAccordion allowMultiple>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.start_heading}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.Understand_pera2} {translations?.ad_p1}
                          </li>
                          <li>{translations?.ad_p2}:</li>
                          <li className="pl-6">
                            {translations?.Iterative_pera2}: "{translations?.ad_p3}"
                          </li>
                          <li className="pl-6">
                            {translations?.ad_p4}: "{translations?.ad_p5}"
                          </li>
                          <li className="pl-6">
                            {translations?.ad_p6}: "{translations?.ad_p7}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Use_Multiple_heading}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.ad_p9}
                          </li>
                          <li>
                            {translations?.Understand_pera2} {translations?.ad_p10}{" "}
                          </li>
                          <li>{translations?.ad_p12}:</li>
                          <li className="pl-6">
                            "{translations?.ad_p13}"
                          </li>
                          <li className="pl-6">
                            "{translations?.ad_p14}"
                          </li>
                          <li className="pl-6">
                            "{translations?.ad_p15}"
                          </li>
                          <li className="pl-6">
                            "{translations?.ad_p16}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Refine_with_Follow_heading}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.ad_p17}
                          </li>
                          <li>{translations?.Understand_pera2} </li>
                          <li className="pl-6">
                            {translations?.Iterative_pera2}: "{translations?.ad_p18}"
                          </li>
                          <li className="pl-6">
                            {translations?.ad_p19}: "{translations?.ad_p20}"
                          </li>
                          <li className="pl-6">
                            {translations?.ad_p6}: "{translations?.ad_p22}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Ask_for_Summaries_heading}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.ad_p23}
                          </li>
                          <li>
                            {translations?.Understand_pera2} {translations?.ad_p24}:{" "}
                          </li>
                          <li className="pl-6">
                            {translations?.Iterative_pera2}: "{translations?.ad_p25}"
                          </li>
                          <li className="pl-6">
                            {translations?.ad_p26}: "{translations?.ad_p27}"
                          </li>
                          <li className="pl-6">
                            {translations?.ad_p28}: "{translations?.ad_p29}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Layer_for_Creative_heading}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.ad_p30}
                          </li>
                          <li>{translations?.Understand_pera2} {translations?.ad_p31} </li>
                          <li className="pl-6">
                            "{translations?.ad_p32}"
                          </li>
                          <li className="pl-6">
                            {translations?.ad_p33}: "{translations?.ad_p34}"
                          </li>
                          <li className="pl-6">
                            {translations?.ad_p6}: "{translations?.ad_p35}""
                          </li>
                        </div>
                      </SdfExpandableBox>
                      <SdfExpandableBox
                        opened={isPromptOpen}
                        header={translations?.Ask_for_Different_heading}
                      >
                        <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                          <li>
                            {translations?.ad_p36}
                          </li>
                          <li>
                            {translations?.Understand_pera2} "{translations?.ad_p37}"{" "}
                          </li>
                          <li className="pl-6">
                            {translations?.ad_p33}: "{translations?.ad_p38}""
                          </li>
                          <li className="pl-6">
                            {translations?.ad_p6}: "{translations?.ad_p39}"
                          </li>
                        </div>
                      </SdfExpandableBox>
                    </SdfAccordion>
                  </>
                )}
              </div>
            ) : (
              <div className="p-2">
                <div className="flex justify-end pt-2 pb-2">
                  <SdfButton
                    emphasis="tertiary"
                    onClick={() => {
                      setIsPromptOpen(!isPromptOpen);
                    }}
                  >
                    {isPromptOpen ? translations?.Collapse_All : translations?.Expand_All}
                  </SdfButton>
                </div>
                <SdfAccordion allowMultiple>
                  <SdfExpandableBox
                    opened={isPromptOpen}
                    header={translations?.General_Functionality}
                  >
                    <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                      <li>
                        <b>{translations?.General_Functionality_Que}</b>
                      </li>
                      <p className="pl-9">
                        {translations?.General_Functionality_pera1}
                      </p>

                      <p className="pl-9">
                        {translations?.General_Functionality_pera2}
                      </p>
                      <li>
                        <b>
                          {translations?.General_Functionality_Que2}
                        </b>
                      </li>
                      <p className="pl-9">
                        {translations?.General_Functionality_ans2}
                      </p>
                    </div>
                  </SdfExpandableBox>
                  <SdfExpandableBox
                    opened={isPromptOpen}
                    header={translations?.Usage_and_Access}
                  >
                    <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                      <li>
                        <b>
                          {translations?.Usage_Que}
                        </b>
                      </li>
                      <p className="pl-9">
                        {translations?.Usage_pera1}
                      </p>
                      <li>
                        <b>
                          {translations?.Usage_Que2}
                        </b>
                      </li>
                      <p className="pl-9">
                        {translations?.Usage_pera2}
                      </p>
                      <li>
                        <b>{translations?.Usage_Que3}</b>
                      </li>
                      <p className="pl-9">
                        {translations?.Usage_pera3}
                      </p>

                    </div>
                  </SdfExpandableBox>
                  <SdfExpandableBox
                    opened={isPromptOpen}
                    header={translations?.Features_and_Capabilities}
                  >
                    <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                      <li>
                        <b>
                          {translations?.Features_Que}
                        </b>
                      </li>
                      <p className="pl-9">
                        {translations?.Features_pera1}
                      </p>

                      <li>
                        <b>
                          {translations?.Features_Que2}
                        </b>
                      </li>
                      <p className="pl-9">
                        {translations?.Features_pera2}
                      </p>
                      <li>
                        <b>
                          {translations?.Features_Que3}
                        </b>
                      </li>
                      <p className="pl-9">
                        {translations?.Features_pera3}
                      </p>
                      <p className="pl-9">
                        {translations?.Features_pera4}
                      </p>
                      <p className="pl-9">
                        {translations?.Features_pera5}
                      </p>
                    </div>
                  </SdfExpandableBox>

                  <SdfExpandableBox
                    opened={isPromptOpen}
                    header={translations?.Troubleshooting_and_Support_Functionality}
                  >
                    <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                      <li>
                        <b>
                          {translations?.Troubleshooting_Que}
                        </b>
                      </li>
                      <p className="pl-9">
                        {translations?.Troubleshooting_pera}
                      </p>

                    </div>
                  </SdfExpandableBox>
                  <SdfExpandableBox
                    opened={isPromptOpen}
                    header={translations?.Security_and_Privacy}
                  >
                    <div className="Prompts  pl-9 pt-5 flex flex-col gap-2">
                      <li>
                        <b>
                          {translations?.Security_Que}
                        </b>
                      </li>
                      <p className="pl-9">
                        {translations?.Security_pera1}
                      </p>
                      <li>
                        <b>
                          {translations?.Security_Que2}
                        </b>
                      </li>
                      <p className="pl-9">
                        {translations?.Security_pera2}
                      </p>
                      <li>
                        <b>
                          {translations?.Security_Que3}
                        </b>
                      </li>
                      <p className="pl-9">
                        {translations?.Security_pera3}{" "}
                      </p>
                    </div>
                  </SdfExpandableBox>
                </SdfAccordion>
              </div>
            )}
          </>
        </SdfFocusPane>

        <SdfFocusPane
          size="large"
          // heading="Acknowledgement"
          hideAcceptButton={true}
          hideDismissButton={true}
          visible={azureUser?.isnoticeRead ? !azureUser.isnoticeRead : false}
          closeable={true}
          onSdfDismiss={(e) => {
            fnuserNoticeAcknowledgment(e, azureUser.opt_in, true);
          }}
        >
          <img
            className="pl-5 pt-0 pb-2"
            src={ADPLogo}
            height={120}
            width={120}
            alt="ADP Logo"
          ></img>

          <div className=" flex justify-center text-xl font-bold pb-4">
            {" "}
            Instructions for Responsible Use of ESI Fusion (Lasted Updated 30th Jan 2026)
          </div>
          <div className="p-4 pt-0 flex flex-col gap-1.5">
            <div>
              ESI Fusion was designed to put the power of an out-of-the-box Large
              Language Model (LLM) functionality in our associates’ hands in a
              way that facilitates their work and transforms standard work tasks
              while protecting our data and our clients’ data. This solution
              integrates an ADP OneUX interface with Azure OpenAI via ADP AI
              Gateway, leveraging AI Guardrails and Azure Personal Identifiable
              Information redaction services.
            </div>
            <div>
              ESI Fusion was designed with several technical guardrails to protect
              our data and information in a way that distinguishes ESI Fusion from
              other third-party Generative AI (GenAI) tools. Data entered into
              ESI Fusion does not train an LLM, and the LLM will not learn from
              ADP associates’ use of ESI Fusion. This is different than other AI
              tools like ChatGPT, which require prior approval from the Chief
              Data Office (CDO). Information is also protected internally; one
              associate’s “conversation”, (which include the information and
              documents input by an associate, and the output received from ESI Fusion
              ) will not become available to other associates using ESI Fusion
              other than for monitoring purposes. ESI Fusion also has several
              technical guardrails designed to protect the personally
              identifiable information (PII) of our associates, clients and
              partners
            </div>
            <div>
              In addition to these technical protections, and because ESI Fusion
              remains in an expanded pilot phase, associates are expected to use
              ESI Fusion for work purposes only, and in a consistent manner with
              ADP’s company policies. In addition, all associates using ESI Fusion
              must adhere to the following guardrails:
            </div>

            <div className="pl-8">
              <li>
                No Sensitive Personal Information. Given the importance of data
                protection, and despite the built-in safeguards, we ask that you
                not input any sensitive personally identifiable information
                (SPI), including bank account or credit card numbers,
                credentials, employment background reports, government-issued ID
                numbers, passwords, PIN codes or protected health information
                into ESI Fusion.
              </li>
              <li>
                No Source Code. ESI Fusion should not be used to develop, modify
                or review source code. Developers should instead use approved
                tools designed for code generation, such as AmazonQ and GitHub
                CoPilot.
              </li>
              <li>
                No Prohibited or High-Risk Use. ESI Fusion may not be used for
                prohibited or high-risk purposes that may impact individual
                rights. AI, including ESI Fusion, should never be used to
                determine a person’s emotional state. AI, including ESI Fusion,
                should not be used to make decisions about hiring, performance
                evaluations, promotions, terminations or work allocation.{" "}
              </li>
            </div>

            <div>
              ESI Fusion relies on developing technology and its output may have
              errors. As with any GenAI tool, associates should review any
              output generated by ESI Fusion for accuracy and potential biases
              before relying on or using it. ESI Fusion output should not replace
              associate discretion or responsibility for making appropriate,
              informed decisions.
            </div>
            <div>
              Associate use of ESI Fusion will be audited to ensure adherence to
              these guardrails. Please raise a SolveIt Ticket if you have
              questions or concerns about the appropriate usage of ESI Fusion.
            </div>
            <div className="pt-5">Copyright © 2025 ADP, Inc.</div>
          </div>
        </SdfFocusPane>

        {messageReadyNotificationContext.is_message_ready &&
          messageReadyNotificationContext?.title && (
            <div
              style={{
                position: "fixed",
                right: "1.30rem",
                top: "8rem",
                "z-index": "50",
                width: "fit-content",
                "max-width": "32rem",
              }}
            >
              <SdfAlertBanner
                size="lg"
                useAnimation={true}
                status="info"
                closeable="true"
                autoClose="true"
                autoCloseAfter="5000"
                onSdfAfterOpen={() =>
                  setTimeout(
                    () =>
                      setMessageReadyNotificationContext({
                        ...messageReadyNotificationContext,
                        is_message_ready: false,
                        // chat_title: "",
                        // chat_id: ""
                      }),
                    5000
                  )
                }
                onSdfAfterClose={() => {
                  setMessageReadyNotificationContext({
                    ...messageReadyNotificationContext,
                    is_message_ready: false,
                    // chat_title: "",
                    // chat_id: ""
                  });
                }}
              >
                <div>
                  <span className="font-bold text-xl">
                    {messageReadyNotificationContext?.title}
                  </span>
                  <div
                    className="cursor-pointer"
                    onClick={(e) =>
                      chatNavigationHandler(e, messageReadyNotificationContext)
                    }
                  >
                    {translations?.Response_received}
                  </div>
                </div>
              </SdfAlertBanner>
            </div>
          )}
        {dispError.status && (
          <div
            style={{
              position: "fixed",
              right: "1.30rem",
              top: "8rem",
              "z-index": "50",
              width: "fit-content",
              "max-width": "32rem",
            }}
          >
            <SdfAlertBanner
              size="xs"
              useAnimation={true}
              status="error"
              icon="alert-error"
              closeable="false"
              autoClose="true"
              autoCloseAfter="5000"
              onSdfAfterOpen={() =>
                setTimeout(
                  () => setdispError({ status: false, message: "" }),
                  5000
                )
              }
              onSdfAfterClose={() => {
                setdispError({ status: false, message: "" });
              }}
            >
              <span
                style={{
                  fontSize: "1rem",
                  lineHeight: "1.75rem",
                  whiteSpace: "pre-wrap",
                }}
              >
                {dispError.message}
              </span>
            </SdfAlertBanner>
          </div>
        )}

        {dispAlertMessage.status && (
          <div
            style={{
              position: "fixed",
              right: "1.30rem",
              top: "5rem",
              "z-index": "9999",
              width: "fit-content",
              "max-width": "32rem",
            }}
          >
            <SdfAlertBanner
              size="xs"
              useAnimation={true}
              status={dispAlertMessage.code}
              icon={
                dispAlertMessage.code === "success"
                  ? "alert-success"
                  : "alert-error"
              }
              closeable={true}
              autoClose={true}
              autoCloseAfter={5000}
              onSdfDismiss={() =>
                setdispAlertMessage({ status: false, code: "", message: "" })
              }
              onSdfAfterOpen={() =>
                setTimeout(
                  () =>
                    setdispAlertMessage({
                      status: false,
                      code: "",
                      message: "",
                    }),
                  5000
                )
              }
              // onSdfAfterClose = {(e) => console.log(e)}
              onSdfAfterClose={(e) => {
                console.log("event after close triggered", e);
                setdispAlertMessage({ status: false, code: "", message: "" });
              }}
            >
              <span
                style={{
                  fontSize: "1rem",
                  lineHeight: "1.75rem",
                  whiteSpace: "pre-wrap",
                }}
              >
                {dispAlertMessage.message}
              </span>
            </SdfAlertBanner>
          </div>
        )}

        {isLoading ? (
          <div
            key={new Date()}
            className="fixed top-0 left-0 w-full h-full zIndexval flex justify-center flex-col gap-4 items-center bg-neutral-100 bg-opacity-20"
            style={{ backdropFilter: "blur(2px)" }}
          >
            <SdfBusyIndicator size="md" className="absolute" />
            <span className="text-lg text-zinc-600 font-semibold mt-24">
               {translations?.Loading}
            </span>
          </div>
        ) : (
          <></>
        )}

        {isDownloading ? (
          <div
            key={new Date()}
            className="fixed top-0 left-0 w-full h-full zIndexval flex justify-center flex-col gap-4 items-center bg-neutral-100 bg-opacity-20"
            style={{ backdropFilter: "blur(2px)" }}
          >
            <SdfBusyIndicator size="md" className="absolute" />
            <span className="text-lg text-zinc-600 font-semibold mt-24">
              {translations?.Downloading}
            </span>
          </div>
        ) : (
          <></>
        )}
        {uploadHanledClick ? (
          <UploadGeneral
            setuploadHanledClick={setuploadHanledClick}
            fnHandleUpload={fnHandleMultiUpload}
            document_id={document_id}
            chatId={chatID}
            requestsession={localStorage.getItem("chatsessionid")}
          />
        ) : (
          <></>
        )}


        <SdfFocusPane
          size="small"
          heading={translations?.Enter_Folder_Title}
          acceptLabel={translations?.Submit}
          hideDismissButton={true}
          visible={isupdateFolderClicked}
          closeable={true}
          // status="info"
          disableAcceptButton={
            folderTitle.replaceAll(" ", "").length === 0 ? true : false
          }
          onSdfDismiss={() => {
            setIsupdateFolderClicked(false);
            setChangeTitleError(false);
            inputQueryRef.current.focus();
          }}
          onSdfAccept={() => {
            setIsupdateFolderClicked(false);
            fnHandleUpdateFolder();
            inputQueryRef.current.focus();
          }}
          hideAcceptButton={folderTitle?.length > 30 ? true : false}
        >
          <div className="flex flex-col p-3 gap-2">
            <SdfTextarea
              ref={editfoldertitleRef}
              className="mt-3"
              rows={2}
              value={folderTitle}
              resize="none"
              state={changeTitleError ? "error" : "normal"}
              maxlength={31}
              // onKeyDown={(e) =>{
              //   console.log(e)
              // }}
              onKeyDown={(e) => {
                if (
                  e.key === "Enter" ||
                  (e?.shiftKey && e.key === "Enter")
                ) {
                  e.preventDefault();
                }
              }}
              onSdfInput={(e) => {
                if (e.target.value.length > 30) {
                  setChangeTitleError(true);
                } else {
                  console.log(e.target.value.length);
                  setChangeTitleError(false);
                }
                setFolderTitle(e.target.value);
              }}
            ></SdfTextarea>
            <span className="text-xs">
              {" "}
              {translations?.update_the_folder_name}
            </span>
            {changeTitleError && (
              <SdfAlertInline status="error">
                {translations?.Title_can_not_exceed}
              </SdfAlertInline>
            )}
          </div>
        </SdfFocusPane>
        <SdfFocusPane
          size="small"
          heading={translations?.Enter_Folder_Title}
          acceptLabel={translations?.Submit}
          hideDismissButton={true}
          visible={isAddFolderClicked}
          closeable={true}
          // status="info"
          disableAcceptButton={
            folderTitle.replaceAll(" ", "").length === 0 ? true : false
          }
          onSdfDismiss={() => {
            SetIsAddFolderClicked(false);
            setChangeTitleError(false);
            setTimeout(() => inputQueryRef.current.focus(), 400);
          }}
          onSdfAccept={() => {
            SetIsAddFolderClicked(false);
            fnHandleAddFolder();
            setTimeout(() => inputQueryRef.current.focus(), 400);
          }}
          hideAcceptButton={folderTitle?.length > 30 ? true : false}
        >
          <div className="flex flex-col p-3 gap-2">
            <SdfTextarea
              ref={foldertitleRef}
              className="mt-3"
              rows={2}
              value={folderTitle}
              resize="none"
              state={changeTitleError ? "error" : "normal"}
              maxlength={31}
              // onKeyDown={(e) =>{
              //   console.log(e)
              // }}
              onKeyDown={(e) => {
                if (
                  e.key === "Enter" ||
                  (e?.shiftKey && e.key === "Enter")
                ) {
                  e.preventDefault();
                }
              }}
              onSdfInput={(e) => {
                if (e.target.value.length > 30) {
                  setChangeTitleError(true);
                } else {
                  console.log(e.target.value.length);
                  setChangeTitleError(false);
                }
                setFolderTitle(e.target.value);
              }}
            ></SdfTextarea>
            <div>
              <div className="text-xs mt-2">{translations?.Whats_a_Folder}</div>
              <span className="text-xs ml-2">
                {translations?.Group_chats_by_client}
              </span>
              <span className="text-xs ml-2">
                {translations?.Less_clutter}
              </span>
            </div>
            {changeTitleError && (
              <SdfAlertInline status="error">
                {translations?.Title_can_not_exceed}
              </SdfAlertInline>
            )}
          </div>
        </SdfFocusPane>
        <SdfFocusPane
          size="small"
          heading={translations?.new_title}
          acceptLabel={translations?.Submit}
          hideDismissButton={true}
          visible={isChangeClicked}
          closeable={true}
          // status="info"
          disableAcceptButton={
            newTitle.replaceAll(" ", "").length === 0 ? true : false
          }
          onSdfDismiss={() => {
            setisChangeClicked(false);
            setChangeTitleError(false);
            inputQueryRef.current.focus();
          }}
          onSdfAccept={() => {
            setisChangeClicked(false);
            fnChangeTitle();
            inputQueryRef.current.focus();
          }}
          hideAcceptButton={newTitle?.length > 30 ? true : false}
        >
          <div className="flex flex-col p-3 gap-2">
            <SdfTextarea
              ref={titleRef}
              className="mt-3"
              rows={2}
              value={newTitle}
              resize="none"
              state={changeTitleError ? "error" : "normal"}
              maxlength={31}
              // onKeyDown={(e) =>{
              //   console.log(e)
              // }}
              onKeyDown={(e) => {
                if (
                  e.key === "Enter" ||
                  (e?.shiftKey && e.key === "Enter")
                ) {
                  e.preventDefault();
                }
              }}
              onSdfInput={(e) => {
                if (e.target.value.length > 30) {
                  setChangeTitleError(true);
                } else {
                  console.log(e.target.value.length);
                  setChangeTitleError(false);
                }
                setnewTitle(e.target.value);
              }}
            ></SdfTextarea>
            <span className="text-xs">
              {" "}
              {translations?.update_chat_name}
            </span>
            {changeTitleError && (
              <SdfAlertInline status="error">
                {translations?.Title_can_not_exceed}
              </SdfAlertInline>
            )}
          </div>
        </SdfFocusPane>

        <SdfFocusPane
          size="small"
          heading={translations?.Confirm_Delete}
          acceptLabel={translations?.Delete}
          hideDismissButton={true}
          visible={isDeleteClicked}
          closeable={true}
          // status="info"
          onSdfDismiss={() => {
            setisDeleteClicked(false);
          }}
          onSdfAccept={() => {
            setisDeleteClicked(false);
            fnRemoveChat();
          }}
          hideAcceptButton={newTitle?.length > 30 ? true : false}
        >
          <div className="flex flex-col p-3 gap-2">
            <span>{translations?.Delete_chat_message}</span>
          </div>
        </SdfFocusPane>
        <SdfFocusPane
          size="small"
          heading={translations?.Confirm_Delete}
          acceptLabel={translations?.Delete}
          hideDismissButton={true}
          visible={isDeleteFolderClicked}
          closeable={true}
          // status="info"
          onSdfDismiss={() => {
            setIsDeleteFolderClicked(false);
            setDeletedFolderId();
          }}
          onSdfAccept={() => {
            setIsDeleteFolderClicked(false);
            fnHandleDeleteFolder();
          }}
        >
          <div className="flex flex-col p-3 gap-2">
            <span>
              zzzzzzzzzzzz
              {translations?.Associated_chat_delete}
            </span>
          </div>
        </SdfFocusPane>
        <SdfFocusPane
          size="small"
          heading={translations?.Confirm_Delete}
          acceptLabel={translations?.Delete}
          hideDismissButton={true}
          visible={isDeleteFolderClicked}
          closeable={true}
          // status="info"
          onSdfDismiss={() => {
            setIsDeleteFolderClicked(false);
            setDeletedFolderId();
          }}
          onSdfAccept={() => {
            setIsDeleteFolderClicked(false);
            fnHandleDeleteFolder();
          }}
        >
          <div className="flex flex-col p-3 gap-2">
            <span>
              {translations?.Associated_chat_delete}
            </span>
          </div>
        </SdfFocusPane>

        <SdfFocusPane
          icon="media-folder"
          size="small"
          heading={translations?.Move_to_Folder}
          acceptLabel={translations?.Submit}
          hideDismissButton={true}
          visible={
            chatsSectionNavstate.isNavActionClicked &&
            chatsSectionNavstate.action === "Move"
          }
          closeable={true}
          // status="info"
          disableAcceptButton={
            chatsSectionNavstate?.Target_folder_id?.replaceAll(" ", "")
              .length === 0
              ? true
              : false
          }
          onSdfDismiss={() => {
            setChatsSectionNavState((prev) => {
              return {
                ...prev,
                isNavActionClicked: false,
                action: "",
                context: "",
                Target_folder_id: "",
                Target_folder_label: "",
              };
            });

            inputQueryRef.current.focus();
          }}
          onSdfAccept={() => {
            fnMoveChat(
              chatsSectionNavstate.chat_id,
              chatsSectionNavstate.Target_folder_id
            );
            setChatsSectionNavState((prev) => {
              return {
                ...prev,
                isNavActionClicked: false,
                action: "",
                context: "",
                Target_folder_id: "",
                Target_folder_label: "",
              };
            });
            inputQueryRef.current.focus();
          }}
        >
          <div className="flex flex-col p-3 gap-2">
            <SdfSelectSimple

              label={`${translations?.Select_a_folder}..`}
              filterable
              portalEnabled
              items={chatsSectionNavstate.context}
              value={{
                label: chatsSectionNavstate.Target_folder_label,
                value: chatsSectionNavstate.Target_folder_label,
              }}
              onSdfChange={(e) =>
                setChatsSectionNavState((prev) => {
                  return {
                    ...prev,
                    Target_folder_id: e.detail.value,
                    Target_folder_label: e.detail.label,
                  };
                })
              }
            ></SdfSelectSimple>
            <span className="text-xs">
              {translations?.Organize_your_chats}
            </span>
            {/* <Span></Span> */}
          </div>
        </SdfFocusPane>

        <SdfFocusPane
          size="small"
          heading={translations.Confirm_remove}
          acceptLabel={translations.Remove}
          hideDismissButton={true}
          visible={
            chatsSectionNavstate.isNavActionClicked &&
            chatsSectionNavstate.action === "Remove"
          }
          closeable={true}
          // status="info"
          onSdfDismiss={() => {
            setChatsSectionNavState((prev) => {
              return {
                ...prev,
                isNavActionClicked: false,
                action: "",
                context: "",
                Target_folder_id: "",
                Target_folder_label: "",
              };
            });
          }}
          onSdfAccept={() => {
            fnMoveChat(
              chatsSectionNavstate.chat_id,
              chatsSectionNavstate.Target_folder_id
            );
            setChatsSectionNavState((prev) => {
              return {
                ...prev,
                isNavActionClicked: false,
                action: "",
                context: "",
                Target_folder_id: "",
                Target_folder_label: "",
              };
            });
            inputQueryRef.current.focus();
          }}
        >
          <div className="flex flex-col p-3 gap-2">
            <span>
              {translations.Chat_move_confirmation}
            </span>
          </div>
        </SdfFocusPane>

{!newOpenGPTScreen && (
  <>
    {hasTransitionedIn || isMounted ? (
      <div className="chatLeftOfWindowsp3">
        {selectedPrompt === "ekm4" && (
          <div className="m-4">
            <FilterDropdownMenu 
              onFilterChange={(updatedList) => {
                console.log("Parent received updated filters:", updatedList);
                setActiveKnowledgeFilters(updatedList);
              }} 
            />
          </div>
        )}
      </div>
    ) : (
      <div className="chatLeftOfWindowUnselect"></div>
    )}
  </>
)}

        {
          openAIGIFlow ? <AIGI_doc_extraction_new/>:
           openFIJ ? (
    <div className="chatwindowcenter flex flex-col justify-between">
      <div
        className="chatSpace mt-8 grid grid-cols-12"
        style={{ marginTop: "1.5rem" }}
      >
        <div className="col-start-2 col-end-10">
          <FIJDocumentProcessor
            setOpenFIJ={setOpenFIJ}
            darkModeVars={darkModeVars}
            setdispAlertMessage={setdispAlertMessage}
          />
        </div>
      </div>
    </div>
  ):
          openBPQ ? (
          <div className="chatwindowcenter flex flex-col justify-between">
            <div
              className="chatSpace mt-8 grid grid-cols-12"
              style={{ marginTop: "1.5rem" }}
            >
              <div className="col-start-2 col-end-10">
                <BPQDocumentProcessor
                  setOpenBPQ={setOpenBPQ}
                  darkModeVars={darkModeVars}
                  setdispAlertMessage={setdispAlertMessage}
                />
              </div>
            </div>
          </div>
        ) :
          <>
             {!openGPTScreen && !newOpenGPTScreen ? (
          <>
            {!blocked_chat_context?.is_blocked &&
              !["rk01", "rk02", "rk03", "ekm4","EL1"].includes(selectedPrompt) && (
                <CustomSelect
                  fnNewChatApp={Model_state_rest}
                  openSideNav={openSideNav}
                  selectedPrompt={selectedPrompt}
                  {...{gptMode}}
                ></CustomSelect>
              )}
            <div className="chatwindowcenter flex flex-col justify-between">
              <div
                className="chatSpace mt-1 grid grid-cols-12"
                style={{ marginTop: "1.5rem" }}
              >
                <div className="col-start-2 col-end-10">
                  {!chatID && (showWelcomeMessage || isHome) ? (
                    <SdfBox className="w-full h-fit p-4" variant="clear">
                      <span>
                        <span className="asd" style={{ fontSize: "26px" }}>
                          {translations.greeting}, {azureUser.firstname}
                        </span>
                        <br />
                        {/* {window.devicePixelRatio} */}
                        <span className="asd1" style={{ fontSize: "22px" }}>
                          {translations.greeting_subtext}
                        </span>
                      </span>
                    </SdfBox>
                  ) : (
                    <></>
                  )}
                  {newChatLoading && (
                    <>
                      <SdfBusyIndicator overlay></SdfBusyIndicator>
                    </>
                  )}
                  {loadingFromRecent ? (
                    <>
                      <SdfBusyIndicator overlay></SdfBusyIndicator>
                    </>
                  ) : (
                    <>
                      {isHome ? (
                        <>
                          <div
                            className={
                              windowDimensions.height > 750
                                ? "grid grid-cols-12 homeCards gap-2 p-1"
                                : "grid grid-cols-12 gap-2 p-1"
                            }
                          >
                            <SdfBox
                              variant={darkModeVars.homeBoxes}
                              className="flex flex-col items-start  promoboxAnimate col-span-4  "
                            >
                              <SdfSpotIllustration
                                size="sm"
                                illustrationName="chat"
                              ></SdfSpotIllustration>
                              <p className="text-xl text-xl font-bold block cardsTitle">
                                {translations.Support}
                              </p>
                              <p className="block mt-2 mb-2 grow-1 cardsDescription">
                                {translations.FAQs_subtask}
                              </p>

                              <div
                                className="-mt-1 mt-auto buttoncolor font-semibold  cursor-pointer pt-2 pb-2 "
                                onClick={OpenHelpModal}
                              >
                                <span className="underline">{translations.View}</span>
                              </div>
                            </SdfBox>
                            <Help {...{ helpVis, setHelpVis, azureUser, helpFeedbackRef, helpSupportRef, setdispAlertMessage, darkModeVars }} />
                            <SdfBox
                              variant={darkModeVars.homeBoxes}
                              className="flex flex-col items-start promoboxAnimate  col-span-4 "
                            >
                              <SdfSpotIllustration
                                size="sm"
                                illustrationName="certificate"
                              ></SdfSpotIllustration>
                              <p className="text-xl font-bold block grow-1  cardsTitle ">
                                {translations.ADP_AI_Guardrails}
                              </p>
                              <p className="block mt-2 mb-2 grow-1 cardsDescription">
                                {translations.ADP_AI_Guardrails_subtask}
                              </p>

                              <div
                                className="-mt-1 mt-auto buttoncolor font-semibold  cursor-pointer pt-2 pb-2 "
                                onClick={() => {
                                  window.open(newpdfLimitation, "_blank");
                                }}
                              >
                                <span className="underline">{translations.View}</span>
                              </div>
                            </SdfBox>
                            <SdfBox
                              variant={darkModeVars.homeBoxes}
                              className="flex flex-col items-start promoboxAnimate  col-span-4 "
                            >
                              <SdfSpotIllustration
                                size="sm"
                                illustrationName="teacher"
                              ></SdfSpotIllustration>
                              <p className="text-xl font-bold block  cardsTitle  ">
                                {translations.Prompt_Guide}
                              </p>
                              <p className="block mt-2 mb-2 grow-1 cardsDescription">
                                {translations.Prompt_Guide_subtask}
                              </p>

                              <div
                                className="-mt-1 mt-auto buttoncolor font-semibold  cursor-pointer pt-2 pb-2 "
                                onClick={(event) => {
                                  fnShowDetails(event, translations?.Prompt_Guide);
                                }}
                              >
                                <span className="underline">{translations.View}</span>
                              </div>
                            </SdfBox>
                          </div>
                        </>
                      ) : (
                        <>
                          {console.log("---asdappppppppasd------", isFav)}
                          {chatID && (
                            <div
                              className={`-mr-5 fav-chat flex gap-2 justify-end transition-all duration-500 ease-in-out`}
                            >

                              {folder_id ? (
                                <></>
                              ) : (
                                <SdfTooltip attachmentPoint="right">
                                  <SdfIcon
                                    slot="tooltip-target"
                                    icon={
                                      isFav
                                        ? "feedback-rate-full"
                                        : "feedback-rate-empty"
                                    }
                                    onClick={fnAddFav}
                                    chat_id={chatID}
                                    isFav={isFav}
                                    style={{ fontSize: "1.5rem" }}
                                    className="cursor-pointer btnFeed"
                                  ></SdfIcon>
                                  {isFav
                                    ? translations?.Remove_from_Favourite
                                    : translations?.ADD_to_Favourite}
                                </SdfTooltip>
                              )}
                            </div>
                          )}
                          <div ref={pdfRef}>
                            {responseValues.map((val, index) => {
                              const isUser = val.sender === "user";
                              const stream_check = val?.stream_check ?? null;
                              // const randomKey = `key-${index}-${Date.now()}`;

                              return (
                                <div
                                  className="chat-container m-4"
                                // key={randomKey}
                                >
                                  <div
                                    className={`w-fit relative gap-1 grid grid-flow-row ${isUser ? "ml-auto p-1 max-w-[77vh]" : ""
                                      }`}
                                  >
                                    <div className="flex gap-1 p-2 justify-start items-start">
                                      {!isUser && (
                                        <div className="grid items-start">
                                          {chatimage ? (
                                            <SdfAvatar
                                              shape="circle"
                                              size="xs"
                                              className="box1"
                                              icon="brand-adp-assist-logo"
                                            />
                                          ) : (
                                            <div className="avatarBOT">
                                              <img
                                                width="20"
                                                height="20"
                                                src={brandassitsvg}
                                                alt="BOT"
                                              />
                                            </div>
                                          )}
                                        </div>
                                      )}

                                      <div className="flex flex-col justify-end ">
                                        {isUser && (
                                          <div className="max-w-[88vh] flex flex-col flex-wrap justify-end gap-1 mb-0 fileDisplay-alignment ml-auto">
                                            {documentStatusList.map(
                                              (f) =>
                                                Number(f.message_id) ===
                                                index && (
                                                  <FileDisplayCard
                                                    key={f.key}
                                                    fileObj={f}
                                                    type="chat_prompt"
                                                    error={
                                                      fileUploadComponentErrors.status
                                                    }
                                                    allowDelete={selectedPrompt==='EL1'}
                                                    onRemove={
                                                        selectedPrompt === "EL1"
                                                          ? () => {
                                                              if (f?.state !== "pending") {
                                                                const file_to_remove = documentStatusList.filter(
                                                                  (item) => item.key && item.key === f.key
                                                                );
                                                                Handle_Remove_SuccessFiles(file_to_remove);
                                                              } else {
                                                                Handle_Remove_inlineFile(f);
                                                              }
                                                            }
                                                          : undefined
                                                      }
                                                  />
                                                )
                                            )}
                                          </div>
                                        )}

                                        <div
                                          className={`flex ${isUser
                                              ? "justify-end"
                                              : "justify-start"
                                            } items-start gap-2 `}
                                        >
                                          <div
                                            className={`${chatimage ? "p-2" : "pl-2 pr-2"
                                              } ${isUser
                                                ? "max-w-full chatbubble chatboxright chatbubble-colorright"
                                                : "w-full"
                                              }`}
                                            variant="shadowed"
                                          >
                                            {console.log(
                                              "val.content--------",
                                              val
                                            )}
                                            {val === "..." ? (
                                              <span className="shimmer">
                                                <DynamicStatusMessage
                                                  type={
                                                    selectedPrompt === "ekm4"
                                                      ? "EKM"
                                                      : browerSearchId ===
                                                        "webs"
                                                        ? "WEBS"
                                                        : "loading"
                                                  }
                                                />
                                              </span>
                                            ) : (
                                              <>
                                                {isUser ? (
                                                  <div className="grid grid-cols ml-1 gap-1">
                                                    <Markdown
                                                      className="react-markdown"
                                                      rehypePlugins={rehypeRaw}
                                                      remarkPlugins={remarkGfm}
                                                    >
                                                      {val.isTranslatable ? translations[val.contentKey] : val.content}
                                                    </Markdown>
                                                    <div className="flex items-end justify-end pb-0">
                                                      <span className="timestamp">
                                                        {val.newTimeStamp}
                                                      </span>
                                                    </div>
                                                  </div>
                                                ) : (
                                                  <>
                                                    {val.job_details?.job_id &&
                                                      val.job_details?.status ===
                                                      "in_progress" ? (
                                                      <div className="text-xl font-semibold animate-pulse">
                                                        {
                                                          val.job_details
                                                            .phaseTracker
                                                        }
                                                      </div>
                                                    ) : (
                                                      <div className="relative group snippet-container">
                                                        <MarkDownReply
                                                          className={
                                                            stream_check
                                                              ? "react-markdown  shimmer"
                                                              : "react-markdown"
                                                          }
                                                          rehypePlugins={
                                                            rehypeRaw
                                                          }
                                                          remarkPlugins={
                                                            remarkGfm
                                                          }
                                                        >
                                                          {val.isTranslatable ? (translations[val.contentKey]) : val.content}
                                                        </MarkDownReply>
                                                      </div>
                                                    )}
                                                  </>
                                                )}
                                              </>
                                            )}
                                            {!isUser &&
                                              val !== "..." &&
                                              val?.webex_groups?.data?.length >
                                              0 &&
                                              index + 1 ===
                                              responseValues.length && (
                                                <SdfButton
                                                  className="p-0 "
                                                  emphasis="tertiary"
                                                  onClick={() => {
                                                    setSessionGroupsModal({
                                                      visible: true,
                                                      groups:
                                                        val.webex_groups.data,
                                                      text: val.webex_groups
                                                        .text,
                                                      actions:
                                                        val.webex_groups
                                                          .actions,
                                                    });
                                                  }}
                                                >
                                                  {translations.View_more}...
                                                </SdfButton>
                                              )}
                                            {!isUser &&
                                              val !== "..." &&
                                              val?.webex_groups?.actions
                                                ?.length > 0 &&
                                              index + 1 ===
                                              responseValues.length && (
                                                <div className="flex justify-start gap-1 mt-2">
                                                  {val?.webex_groups?.actions?.map(
                                                    (action, i) => {
                                                      console.log(action);
                                                      return (
                                                        <SdfTag
                                                          readonly={true}
                                                          onClick={() => {
                                                            fnHandleSubmit(
                                                              localStorage.getItem(
                                                                "chatsessionid"
                                                              ),
                                                              selectedPrompt,
                                                              "Webex_Groups_Session",
                                                              action
                                                            );
                                                          }}
                                                          id={i}
                                                          className="cursor-pointer"
                                                          value={action}
                                                        ></SdfTag>
                                                      );
                                                    }
                                                  )}
                                                </div>
                                              )}

                                            {!isUser &&
                                              val !== "..." &&
                                              val?.source_url?.length > 0 && (
                                                <div className="grid grid-cols-12 mt-4 gap-2">
                                                  <SdfBox
                                                    variant="shadowed"
                                                    className="flex flex-col items-start  promoboxAnimate col-span-12  "
                                                  >
                                                    <p className="text-black text-xl block cardsTitle">
                                                      {translations.Source}
                                                    </p>
                                                    {val?.source_url?.map(
                                                      (action) => {
                                                        console.log(action);
                                                        return (
                                                          <>
                                                            <p
                                                              className="text-black  block mt-1 mb-1 grow-1 cardsDescriptionForSource"
                                                              onClick={() => {
                                                                fnactionOpenUrl(
                                                                  action.url
                                                                );
                                                              }}
                                                            >
                                                              <sdf-icon icon="action-link"></sdf-icon>{" "}
                                                              &nbsp;{" "}
                                                              <span className="underline">
                                                                {action.title}
                                                              </span>
                                                            </p>
                                                          </>
                                                        );
                                                      }
                                                    )}
                                                  </SdfBox>
                                                </div>
                                              )}
                                            {!isUser &&
                                              val !== "..." &&
                                              val?.service_used ===
                                              "web search" && (
                                                <div className="flex items-center mt-4 fontSizeMsg websdisclaimerdisplay">
                                                  {web_browser_svg_resp()}
                                                  <div className="italic items-center websdisclaimerdisplay_margin">
                                                    &nbsp;{translations.websearch_response_feature}
                                                  </div>
                                                </div>
                                              )}                                            {!isUser &&
                                                val !== "..." &&
                                                val?.supporting_documents &&
                                                val.supporting_documents.length > 0 && (
                                                  <ImageChartDisplay
                                                    supportingDocuments={val.supporting_documents}
                                                    chatID={chatID}
                                                    azureUser={azureUser}
                                                    translations={translations}
                                                    setIsDownloading={setIsDownloading}
                                                    setdispError={setdispError}
                                                  />
                                                )}
                                            {!isUser &&
                                              val !== "..." &&
                                              (val.job_details?.job_id ===
                                                null ||
                                                val.job_details?.status ===
                                                "ready") && (
                                                <div className="flex flex-row gap-2 justify-start pt-2">
                                                  {chatimage ? (
                                                    <FeedbackSection
                                                      feedbackRef={feedbackRef}
                                                      inputQueryRef={
                                                        inputQueryRef
                                                      }
                                                      data={val}
                                                      fnsendDislikeFeedback={
                                                        fnsendDislikeFeedback
                                                      }
                                                      index={index}
                                                      fnSendLikeFeedback={
                                                        fnSendLikeFeedback
                                                      }
                                                      feedbackCategories={
                                                        feedbackCategories
                                                      }
                                                      azureUser={azureUser}
                                                      userlocaltimeStamp={
                                                        val.newTimeStamp
                                                      }
                                                      supportingDocuments={
                                                        val?.supporting_documents ??
                                                        []
                                                      }
                                                      job_details={
                                                        val?.job_details
                                                      }
                                                      ChatID={chatID}
                                                      setIsDownloading={
                                                        setIsDownloading
                                                      }
                                                      setdispError={
                                                        setdispError
                                                      }
                                                    />
                                                  ) : (
                                                    <span className="timestamp">
                                                      {val.newTimeStamp}
                                                    </span>
                                                  )}
                                                </div>
                                              )}
                                          </div>

                                          {isUser && (
                                            <div className="avatarUser" style={{ background: darkModeVars.backgroundChatUser }}>
                                              <span>{azureUser.initials}</span>
                                            </div>
                                          )}
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </>
                      )}
                    </>
                  )}
                  <span ref={endOfMessageRef1}></span>
                
                 </div>                
                {isHome &&  
                <div className="col-end-13 mr-5 -mt-4 flex justify-end items-start gap-3">
                  {localStorage.getItem("admin")==="true" && <SdfButton className="whitespace-nowrap" emphasis="tertiary" icon="action-add" iconPlacement="before" size="lg" onClick={()=>{setcreateNoteVis(!createNoteVis)}}>Create Notification</SdfButton>}
                  <SdfButton className="whitespace-nowrap" emphasis="tertiary" icon="media-play" iconPlacement="before" size="lg" onClick={() => { if (openSideNav) setOpenSideNav(!openSideNav) }}>{translations?.start_tour}</SdfButton>
                </div>}
              </div>
              <div className="grid grid-cols-12">

              </div>
              <div className="grid grid-cols-12">
                <div className="col-start-2 col-end-10">
                  {gptMode?.is_gpt && responseValues.length === 0 && (

                    <div className="flex justify-center items-center  overflow-y-auto">
                      <div
                        className="self-center p-2 gap-2 flex flex-wrap justify-center"

                      >
                        {gptMode?.details?.starters?.map(
                          (input) =>
                            input !== "" && (
                              <SdfBox
                                style={{
                                  wordBreak: "break-word",
                                  overflowWrap: "anywhere",
                                }}
                                className={`max-w-56 border-2 cursor-pointer`}
                                variant="clear"
                                onclick={() => fnHandleSubmit(
                                  localStorage.getItem("chatsessionid"),
                                  selectedPrompt,
                                  "Agent Starters",
                                  input
                                )}
                              >
                                {input}
                              </SdfBox>
                            )
                        )}

                      </div>
                    </div>
                  )}
                </div>

                {blocked_chat_context?.is_blocked &&
                  blocked_chat_context?.blocked_message !== "" && (
                    <SdfAlertInline
                      className="col-start-2 col-span-8"
                      status="error"
                      size="sm"
                      autoClose="true"
                      autoCloseAfter="5000"
                      onSdfAfterClose={() => {
                        setBlocked_chat_context({
                          is_blocked: false,
                          blocked_message: "",
                        });
                      }}
                    >
                      {blocked_chat_context?.blocked_message}
                    </SdfAlertInline>
                  )}
                  
                  {selectedPrompt && <div className="z-100000 bg-[#253B8F] text-white h-12 px-5 flex items-center justify-between
             bg-[#253B8F]
             text-white
             h-12
             px-5 mb-[0.5px]">
                            <div className="flex items-center gap-3">
                              <span className="text-lg font-semibold">
                                {selectedPrompt + " Assistance"}
                              </span>

                              <SdfIcon icon="action-help" className="text-white text-lg" />
                            </div>

                            <button
                              type="button"
                              onClick={fnNewChat}
                              className="flex items-center justify-center w-8 h-8 hover:bg-white/10 rounded"
                            >
                              <SdfIcon icon="action-close" />
                            </button>
                          </div>}

                <div
                  className={
                    voiceStore.showPane
                      ? "Container w-full"
                      : !blocked_chat_context.is_blocked &&
                        (isResponseDone || responseValuesLimit)
                        ? "Container w-full"
                        : "disableClass Container w-full"
                  }
                  style={{ background: darkModeVars?.sideBarColor }}
                >
                  
                  <div className="pl-2" style={{ width: "98%" }}>
                    <div className="flex flex-wrap mt-1 gap-2">
                      {/* //&& documentStatusList.filter(file => file.state === 'pending') */}
                      {filesToUpload.length > 0 &&
                        filesToUpload.map((file) => {
                          return (
                            <FileDisplayCard
                              key={file.key}
                              fileObj={file}
                              error={fileUploadComponentErrors.status}
                              type="Inline_prompt"
                              onRemove={() => {
                                if (file?.state !== "pending") {
                                  const file_to_remove =
                                    documentStatusList.filter(
                                      (f) => f.key && f.key === file.key
                                    );
                                  Handle_Remove_SuccessFiles(file_to_remove);
                                } else {
                                  Handle_Remove_inlineFile(file);
                                }
                              }}
                            />
                          );
                        })}
                    </div>

                    {fileUploadComponentErrors.status && (
                      <SdfAlertInline status="error">
                        {translations?.[fileUploadComponentErrors.message] ||
                          fileUploadComponentErrors.fallbackMessage ||
                          "Something went wrong"}
                      </SdfAlertInline>
                    )}

                    {
                      // Only show shimmer if NO pre-validation errors exist
                      !fileUploadComponentErrors?.hasprevalidationError &&
                      documentStatusList.filter(
                        (f) =>
                          f.status !== "Success" &&
                          f.status !== "Failed" &&
                          f.state === "pending"
                      ).length > 0 && (
                        <span className="shimmer">
                          <DynamicStatusMessage
                            type="fileupload"
                            intervalMs={10000}
                          />
                        </span>
                      )
                    }
                    <textarea
                      // disabled={listeningWaveForm}
                      className={
                        listeningWaveForm
                          ? "dynamic-textarea disableClass"
                          : "dynamic-textarea"
                      }
                      style={{ background: darkModeVars?.sideBarColor }}
                      maxlength={8193}
                      ref={inputQueryRef}
                      placeholder={translations.Ask_Anything}
                      id="inputQuery"
                      rows={2}
                      // value={inputQueryRef.current.value}
                      onChange={(event) => {
                        console.log(event);

                        event.target.style.height = "auto";
                        event.target.style.height = `${event.target.scrollHeight}px`;
                      }}
                      onPaste={(e) => {
                        const files = e.clipboardData.files;
                        if (files.length > 0) {
                          e.preventDefault();     
                          const isAttachmentDisabled =
                            azureUser?.selected_model?.document_supported === 'No' ||
                            ["rk01", "rk02", "rk03", "ekm4"].includes(selectedPrompt) ||
                            browerSearchId === "webs"
                          if (isAttachmentDisabled) {
                             setChatInputErrorMessage({
                                status: true,
                                message: `❌ Images and Files are not supported in this model.`
                              });
                              return;
                          }
                          const fileArray = Array.from(files); // ← multiple files supported

                          if (onFilePasteFn) onFilePasteFn(fileArray);
                        }
                      }}
                      onInput={(e) => {
                        if (isResponseDone || responseValuesLimit) {
                          console.log("entered input evenet", e);
                          inputQueryRef.current.value = e.target.value;

                          if (transcript && transcript.trim() !== "") {
                            console.log("dispatch transcript");
                            dispatch(setTranscript(e.target.value));
                          }

                          if (e.target.value.length > 8192) {
                            setChatInputErrorMessage({
                              status: true,
                              message:
                                translations.Message_exceed_Limit
                            });
                          } else {
                            if (chatInputErrorMessage.status) {
                              setChatInputErrorMessage({
                                status: false,
                                message: "",
                              });
                            }
                          }
                        }
                        //setMessage(e.target.value);
                      }}
                      onKeyDown={(e) => {
                        console.log(e);
                        console.log(e.key);
                        if (blocked_chat_context?.is_blocked) {
                          e.preventDefault();
                          return;
                        }
                        if (isResponseDone || responseValuesLimit) {
                          if (!e?.shiftKey && e.key === "Enter") {
                            e.preventDefault();
                            e.target.style.height = "58px";
                            // e.target.style.padding='10px';
                            set_voice_mode("voice_inline");
                            stopListening();
                            dispatch(setStopListening());
                            fnHandleSubmit(
                              localStorage.getItem("chatsessionid"),
                              selectedPrompt
                            );
                            setListeningWaveForm(false);
                          }
                        }
                      }}
                    ></textarea>

                    <div
                      className="flex justify-between"
                      style={{ width: "99%" }}
                    >
                      <div className=" flex flex-row gap-2 justify-between">
                          <div className={`${gptMode?.details?.document_id?.length > 0 && "disableClass"}`}>
                            {/* <FileUploadComponent
                              registerPasteHandler={registerPasteHandler}
                              current_message_Id={responseValues.length}
                              setFileUploadComponentErros={
                                setFileUploadComponentErros
                              }
                              setDocumentStatusList={setDocumentStatusList}
                              setFilesToUpload={setFilesToUpload}
                              onFilesValidated={onFilesValidated}
                              isFileProcessing={isFileProcessing}
                              documentStatusList={documentStatusList}
                              selectedPrompt={selectedPrompt}
                              browerSearchId={browerSearchId}
                            // chatInputErrorMessage={chatInputErrorMessage}
                            ></FileUploadComponent> */}
                            <AttachmentActionMenu
                              registerPasteHandler={registerPasteHandler}
                              current_message_Id={responseValues.length}
                              setFileUploadComponentErros={
                                setFileUploadComponentErros
                              }
                              setDocumentStatusList={setDocumentStatusList}
                              setFilesToUpload={setFilesToUpload}
                              onFilesValidated={onFilesValidated}
                              isFileProcessing={isFileProcessing}
                              documentStatusList={documentStatusList}
                              selectedPrompt={selectedPrompt}
                              browerSearchId={browerSearchId}
                            // chatInputErrorMessage={chatInputErrorMessage}
                            setselectedPrompt={setselectedPrompt}
                            fnClickPrompt={fnClickPrompt}
                            ></AttachmentActionMenu>
                          </div>
                      </div>
                      {voiceStore.isListening ? (
                        <>
                          <VoiceWaveform isListening={voiceStore.isListening} />
                        </>
                      ) : (
                        <></>
                      )}

                      <div className="flex gap-1.5">
                        {
                          <>
                            <InlineVoice
                              {...{ darkModeVars }}
                              inputQueryRef={inputQueryRef}
                              listeningWaveForm={listeningWaveForm}
                              setListeningWaveForm={setListeningWaveForm}
                              set_voice_mode={set_voice_mode}
                              selectedPrompt={selectedPrompt}
                              fnHandleSubmit={fnHandleSubmit}
                              startListening={startListening}
                              stopListening={stopListening}
                              dispatch={dispatch}
                              setDisableMute={setDisableMute}
                              setIsMuted={setIsMuted}
                              setMessage={setMessage}
                              checkMicrophoneAvailable={
                                checkMicrophoneAvailable
                              }
                              setmicrophoneAvailable={setmicrophoneAvailable} // normal version of this variable is not used if everything is ok remove this variable
                              setvoiceDisplayText={setvoiceDisplayText}
                              setInputRefVal={setInputRefVal}
                              transcript={transcript}
                              resetTranscript={resetTranscript}
                              listening={listening}
                              browserSupportsSpeechRecognition={
                                browserSupportsSpeechRecognition
                              }
                            />

                            <VoiceChat
                              setselectedPrompt={setselectedPrompt}
                              set_voice_mode={set_voice_mode}
                              inputQueryRef={inputQueryRef}
                              setvoiceDisplayText={setvoiceDisplayText}
                              setDisableMute={setDisableMute}
                              setIsMuted={setIsMuted}
                              setMessage={setMessage}
                              checkMicrophoneAvailable={
                                checkMicrophoneAvailable
                              }
                              setmicrophoneAvailable={setmicrophoneAvailable}
                              stopListening={stopListening}
                              voiceStore={voiceStore}
                              voiceDisplayText={voiceDisplayText}
                              disableMute={disableMute}
                              isMuted={isMuted}
                              audioRef={audioRef}
                              setListeningWaveForm={setListeningWaveForm}
                              controls={controls}
                              voice_mode={voice_mode}
                              selectedPrompt={selectedPrompt}
                              fnHandleSubmit={fnHandleSubmit}
                              startListening={startListening}
                              transcript={transcript}
                              resetTranscript={resetTranscript}
                              listening={listening}
                              browserSupportsSpeechRecognition={
                                browserSupportsSpeechRecognition
                              }
                              azureUser={azureUser}
                            />
                          </>
                        }
                        <SdfTooltip attachmentPoint="right">
                          <SdfAvatar
                            className={
                              isResponseDone
                                ? "cursor-pointer avatarSend"
                                : "disableClass cursor-pointer avatarSend"
                            }
                            icon="nav-move-up"
                            emphasis="primary"
                            slot="tooltip-target"
                            size="xs"
                            onClick={() => {
                              setTimeout(() => {
                                const textarea =
                                  document.getElementById("inputQuery");
                                textarea.style.height = "58px";
                              }, 500);
                              set_voice_mode("voice_inline");

                              fnHandleSubmit(
                                localStorage.getItem("chatsessionid"),
                                selectedPrompt
                              );
                              stopListening();
                              dispatch(setStopListening());
                              setListeningWaveForm(false);
                            }}
                          ></SdfAvatar>
                          {translations.Send}
                        </SdfTooltip>
                      </div>
                    </div>
                  </div>
                </div>
                  
                {displayErrorMessage.status && message.trim().length === 0 ? (
                  <SdfAlertInline
                    className="col-start-2 col-span-6"
                    status="error"
                    size="sm"
                    autoClose="true"
                    autoCloseAfter="5000"
                    onSdfAfterClose={() => {
                      setdisplayErrorMessage({ status: false, message: "" });
                    }}
                  >
                    {displayErrorMessage.message}
                  </SdfAlertInline>
                ) : (
                  <></>
                )}
                {chatInputErrorMessage.status && (
                  <SdfAlertInline
                    className="col-start-2 col-span-8"
                    status="error"
                    size="sm"
                    autoClose="true"
                    autoCloseAfter="5000"
                    onSdfAfterClose={() => {
                      setChatInputErrorMessage({ status: false, message: "" });
                    }}
                  >
                    {chatInputErrorMessage.message}
                  </SdfAlertInline>
                )}

                <span className="col-start-2 col-end-10 flex flex-col">
                  <span className="text-center fontSizeMsg font-light font-thin mt-1 tctextColor">
                    {translations.footer_Message}
                  </span>
                  <span
                    className="flex justify-center text-l mt-1 font-bold"
                    style={{ color: "rgb(50, 82, 173)" }}
                  >
                    {translations.footer_sub_Message}
                  </span>
                </span>
              </div>
            </div>
          </>
        ) :
          (
            newOpenGPTScreen ? (
              <GPTs {...{ azureUser, fnClickPrompt, setdispAlertMessage, gptData, handleSetGPTdata, myGptDataLoader, setMyGptDataLoader }} />
            ) : (
              <GPTscreen {...{ fnClickPrompt, isEsGroup }} />
            )
          )
        }
          </>
 
  
        }
       


      </div>

      {voiceStore.isError ? (
        <SdfAlertToast
          closeable
          autoClose
          autoCloseAfter={5000}
          status="error"
          useAnimation
          onSdfAfterClose={() => {
            console.log("after close");
            dispatch(setIsError(false));
          }}
          onSdfDismiss={() => {
            console.log("after close 2");
            dispatch(setIsError(false));
          }}
          message={voiceStore.errorMessage}
        ></SdfAlertToast>
      ) : (
        <></>
      )}
      <audio ref={audioRef} src={pollyAudio} />
    </>
  );
};

export default AgentAssist;
