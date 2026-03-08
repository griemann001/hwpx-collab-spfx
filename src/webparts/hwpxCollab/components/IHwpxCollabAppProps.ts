export interface IHwpxCollabAppProps {
  userName: string;
  userEmail: string;
  userLoginName: string;
  userColor: string;
  wsUrl: string;
  spHttpClient: any;  // SPHttpClient - any로 처리하여 타입 충돌 방지
  siteUrl: string;
}
