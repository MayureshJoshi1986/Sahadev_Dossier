$(function(){
	
	function menu_hover(){
		var hovertimer, current;

		$('#nav li.clickable').hover(function(){
			current = $(this).attr('id');
			if( (navigator.platform == 'iPad')) {
				$('#'+current+' .submenu').delay(100).slideDown(100);
			} else {
				hovertimer = setTimeout(function(){ 
					$('#'+current+' .submenu').delay(100).slideDown(100);
				}, 100);
			}
			
		}, function(){
			if( (navigator.platform == 'iPad')) {
				$('.submenu').slideUp(100, function(){
				});
			} else {
				window.clearTimeout(hovertimer);
				setTimeout(function(){
					$('.submenu').slideUp(100, function(){
					});
				}, 100);
			}
		});
		
	}
	menu_hover();
	
	function menu_click(){
		$('#nav li.clickable a.clickable').click(function(){
			if($(this).parent().hasClass('closed')) {
				$('.submenu').slideUp();
				$(this).parent().find('.submenu').slideDown();
			} else if($(this).parent().hasClass('open')) {
				$(this).parent().find('.submenu').slideUp();	
			}
			return false;
		});

	}

	$('.menu-toggle').click(function(){
		$('#nav').slideToggle();		
		$('.submenu').slideUp();	
		$('#nav li').removeClass('open').addClass('closed');	
		return false;
	});
	

	var desktop = false, mobile = false;

	function menu_resize(){
		var window_width = $(window).width();
		if(window_width >= 712){
			if (desktop == false){
				mobile = false;
				desktop = true;
				$('#nav li').removeClass('open');
				$('#nav').show();
				$('#nav li .submenu').hide();
				$('#nav li a').off('click')
				menu_hover();
			}
		} else if(window_width < 712) {
			if (mobile == false) {
				mobile = true;
				desktop = false;
				$('#nav li').off('hover');
				$('#nav li').addClass('closed');
				menu_click();
			}
		}
	}
	
	menu_resize();
	function resize() {
	 	menu_resize();
	}
	var TO = false;
	$(window).resize(function(){
		if(TO !== false)
	   		clearTimeout(TO);
		TO = setTimeout(resize, 200); 
	});
		
});	