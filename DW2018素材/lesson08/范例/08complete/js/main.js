$(function() {
    var timer;
    var iNow = 0;
    fnFade();
    //autoPlay();
    function autoPlay() {
        timer = setInterval(function() {
            iNow++;
            iNow %= $('.bigList li').length;
            fnFade();
        }, 2000)
    }
    $('.smallList li').hover(function() {
        iNow = $(this).index();
        fnFade();
    })
    $('.next').click(function() {
        iNow++;
        iNow %= $('.bigList li').length;
        fnFade();

    })
    $('.pre').click(function() {
        if (iNow == 0) {
            iNow = $('.bigList li').length;
        }
        iNow--;
        fnFade();

    })

    function fnFade() {
        $('.bigList li').each(function(index) {
            if (index != iNow) {
                $('.bigList li').eq(index).fadeOut().css('zIndex', 1);
            } else {
                $('.bigList li').eq(index).fadeIn().css('zIndex', 2);
            }
        })
    }
})