/* Slider */

/* Arrows */
.slick-prev,.slick-next
{
    font-size: 0;
    line-height: 0;
    position:absolute;
    top: 45%;
    display: block;
    width: 25px;
    height: 25px;
    padding: 0;
    cursor: pointer;
    color: transparent;
    border: none;
    outline: none;
	background:#fff;
	border-radius:100%;        /* CSS3草案 */  
    -webkit-border-radius:100%;    /* Safari,Google Chrome用 */  
    -moz-border-radius:100%;   /* Firefox用 */  
}
.slick-next
{
	background-image: url("http://ic4-a.dena.ne.jp/mi/w/1280/h/1280/q/90/bcimg1-a.dena.ne.jp/plus/u3534661/pc/css/../img/bx-next.png");
	background-size:100%;
}
.slick-prev{
	background-image: url("http://ic4-a.dena.ne.jp/mi/w/1280/h/1280/q/90/bcimg1-a.dena.ne.jp/plus/u3534661/pc/css/../img/bx-prev.png");
	background-size:100%;
}
.slick-prev:hover,
.slick-prev:focus,
.slick-next:hover,
.slick-next:focus
{
    color: transparent;
    outline: none;
}
.slick-prev:hover:before,
.slick-prev:focus:before,
.slick-next:hover:before,
.slick-next:focus:before
{
    opacity: 1;
}
.slick-prev.slick-disabled:before,
.slick-next.slick-disabled:before
{
    opacity: .25;
}

.slick-prev:before,
.slick-next:before
{
    background-image: url("http://ic4-a.dena.ne.jp/mi/w/1280/h/1280/q/90/bcimg1-a.dena.ne.jp/plus/u3534661/pc/css/../img/bx-prev.png");
	background-size:90%;
    font-size: 20px;
