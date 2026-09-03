import axios from 'axios';
const api=axios.create({baseURL:'/api'});
api.interceptors.request.use(c=>{const t=localStorage.getItem('sc_token');if(t)c.headers.Authorization=`Bearer ${t}`;return c});
api.interceptors.response.use(r=>r,r=>{if(r.response?.status===401){localStorage.removeItem('sc_token');localStorage.removeItem('sc_user');if(location.pathname!='/login')location.href='/login'}return Promise.reject(r)});
export default api;
