!function(e){var i={};function t(a){if(i[a])return i[a].exports;var n=i[a]={i:a,l:!1,exports:{}};return e[a].call(n.exports,n,n.exports,t),n.l=!0,n.exports}t.m=e,t.c=i,t.d=function(e,i,a){t.o(e,i)||Object.defineProperty(e,i,{enumerable:!0,get:a})},t.r=function(e){"undefined"!=typeof Symbol&&Symbol.toStringTag&&Object.defineProperty(e,Symbol.toStringTag,{value:"Module"}),Object.defineProperty(e,"__esModule",{value:!0})},t.t=function(e,i){if(1&i&&(e=t(e)),8&i)return e;if(4&i&&"object"==typeof e&&e&&e.__esModule)return e;var a=Object.create(null);if(t.r(a),Object.defineProperty(a,"default",{enumerable:!0,value:e}),2&i&&"string"!=typeof e)for(var n in e)t.d(a,n,function(i){return e[i]}.bind(null,n));return a},t.n=function(e){var i=e&&e.__esModule?function(){return e.default}:function(){return e};return t.d(i,"a",i),i},t.o=function(e,i){return Object.prototype.hasOwnProperty.call(e,i)},t.p="",t(t.s=327)}({327:function(e,i,t){"use strict";var a="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(e){return typeof e}:function(e){return e&&"function"==typeof Symbol&&e.constructor===Symbol&&e!==Symbol.prototype?"symbol":typeof e};!function(){var e,i,t,n,s,r,o=/^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;function l(e,i,t,a){$.ajax({type:"POST",url:a+"/login",crossDomain:!0,xhrFields:{withCredentials:!0},data:{email_address:i,password:t},dataType:"json",success:function(a,n,r){s.set("user_email",i),s.set("user_password",t),s.set("user_client_id",e),s.saveAsync(d)},error:function(e){function i(){return e.apply(this,arguments)}return i.toString=function(){return e.toString()},i}(function(){v("error","Login error. Check your credentials and try agaon.",5e3,"#div-authentication-details"),$("#log-message").text(error.responseJSON.detail+", StatusCode: "+error.status+", headers: "+error.getAllResponseHeaders)})})}function d(e){e.status==Office.AsyncResultStatus.Failed||c(Office.context.mailbox.item.from.emailAddress)}function c(t){$(".page-div").addClass("hidden"),$.ajax({type:"GET",url:n+"/search?terms="+t,crossDomain:!0,xhrFields:{withCredentials:!0},data:{email_address:e,password:i},dataType:"json",success:function(e,i,t){if($("#log-message").text("Final success"),e.batch_size>0){var n=!1;$("#log-message").text("Batch size: "+e.batch_size+", typeof: "+a(e.results)),$.each(e.results,function(e,i){if($("#log-message").text("Iterations: Result.type: "+i.type),"salesseek.contacts.models.individual.Individual"==i.type){n=!0,$("#content-organization-name").text(i.item.organization.name),$("#content-organization-role").text(i.item.role);var t="",a="",s="",r="",o="",l="";$.each(i.item.organization.communication,function(e,i){"website"==i.name?(""!=t&&(t+=", "),t+=i.value):"phone"==i.name?(""!=s&&(s+=", "),s+=i.value):"email"==i.name?(""!=a&&(a+=", "),a+=i.value):"linkedin"==i.name?(""!=r&&(r+=", "),r+=i.value):"twitter"==i.name?(""!=o&&(o+=", "),o+=i.value):"facebook"==i.name&&(""!=l&&(l+=", "),l+=i.value)}),$("#content-organization-website").text(t),$("._value_v0gc7_56.email").text(a),$("._value_v0gc7_56.email").attr("href","mailto:"+a),""==a?$("._value_v0gc7_56.email").closest("._FieldContainer_13z8k_1").addClass("hidden"):$("._value_v0gc7_56.email").closest("._FieldContainer_13z8k_1").removeClass("hidden"),$("._value_v0gc7_56.phone").text(s),$("._value_v0gc7_56.phone").attr("href","tel:"+s),""==s?$("._value_v0gc7_56.phone").closest("._FieldContainer_13z8k_1").addClass("hidden"):$("._value_v0gc7_56.phone").closest("._FieldContainer_13z8k_1").removeClass("hidden"),$("#log-message").text("Organization: "+i.item.organization.name),$("._value_v0gc7_56.twitter").text("@"+o),$("._value_v0gc7_56.twitter").attr("href","https://twitter.com/"+o),""==o?$("._value_v0gc7_56.twitter").closest("._FieldContainer_13z8k_1").addClass("hidden"):$("._value_v0gc7_56.twitter").closest("._FieldContainer_13z8k_1").removeClass("hidden"),$("._value_v0gc7_56.facebook").text(l),$("._value_v0gc7_56.facebook").attr("href",l),""==l?$("._value_v0gc7_56.facebook").closest("._FieldContainer_13z8k_1").addClass("hidden"):$("._value_v0gc7_56.facebook").closest("._FieldContainer_13z8k_1").removeClass("hidden"),$("._value_v0gc7_56.linkedin").text(r),$("._value_v0gc7_56.linkedin").attr("href",r),""==r?$("._value_v0gc7_56.linkedin").closest("._FieldContainer_13z8k_1").addClass("hidden"):$("._value_v0gc7_56.linkedin").closest("._FieldContainer_13z8k_1").removeClass("hidden"),$("#div-organization-details").removeClass("hidden"),$("#div-add-user-to-crm").addClass("hidden")}}),n||u()}else u()},error:function(e){function i(){return e.apply(this,arguments)}return i.toString=function(){return e.toString()},i}(function(){$("#log-message").text("Error"),$("#log-message").text(error.responseJSON.detail+", StatusCode: "+error.status+", headers: "+error.getAllResponseHeaders)})})}function u(){if($("#div-organization-details").addClass("hidden"),$("#div-add-user-to-crm").removeClass("hidden"),$.ajax({type:"GET",url:n+"/organizations",crossDomain:!0,xhrFields:{withCredentials:!0},data:{email_address:e,password:i},dataType:"json",success:function(e,i,t){$("#log-message").text("Final success");var a=$("#field-select-organization");$.each(e,function(e,i){a.append('<option value="'+i.id+'">'+i.name+"</option>")})},error:function(e){function i(){return e.apply(this,arguments)}return i.toString=function(){return e.toString()},i}(function(){$("#log-message").text("Error"),$("#log-message").text(error.responseJSON.detail+", StatusCode: "+error.status+", headers: "+error.getAllResponseHeaders)})}),null!=r){var t=r.from.displayName.split(" ");t.length>0&&$(".field-first-name").val(t[0]),t.length>1&&$(".field-last-name").val(t[1])}}function m(){var e=0;return $.each($(".field-name"),function(i,t){""==$.trim($(t).val())?($(this).siblings(".error-message").html($(this).data("name")+" is required."),e++):$(this).siblings(".error-message").html("")}),0==e}function f(){var e=!0;return $("#field-check-organization").is(":checked")?(""==$.trim($("#field-input-organization").val())?(e=!1,$("#field-input-organization").siblings(".error-message").html("New organization name is required.")):$("#field-input-organization").siblings(".error-message").html(""),$("#field-select-organization").siblings(".error-message").html("")):""==$.trim($("#field-select-organization").val())?(e=!1,$("#field-select-organization").siblings(".error-message").html("Organization is required.")):$("#field-select-organization").siblings(".error-message").html(""),e}function g(e){var i=$("#field-first-name").val(),t=$("#field-last-name").val(),a=r.from.emailAddress,s=$("#field-role").val(),o=JSON.stringify({first_name:i,last_name:t,communication:[{name:"Work",medium:"email",value:a,comments:""}],organization_id:e,roles:[{title:s,organization_id:e}]});$.ajax({type:"POST",url:n+"/individuals",crossDomain:!0,xhrFields:{withCredentials:!0},data:o,dataType:"json",success:function(e,i,t){$("#loading-gif-add-individual").addClass("hidden"),$("#loading-gif-alt").removeClass("hidden"),$("#div-add-user-to-crm .detail-section").attr("disabled",!1),$("#div-add-user-to-crm  .alert-message").text(r.from.displayName+" successfully added to SalesSleek."),v("success",r.from.displayName+" successfully added to SalesSleek.",2e3,"#div-add-user-to-crm",c(r.from.emailAddress))},error:function(){$("#loading-gif-add-individual").addClass("hidden"),$("#loading-gif-alt").removeClass("hidden"),$("#div-add-user-to-crm .detail-section").attr("disabled",!1),v("error","Could not add "+r.from.displayName+" to SalesSleek. Please try again or contact system administrator.",5e3,"#div-add-user-to-crm")}})}function v(e,i,t,a,n){"error"==e?$(a+" .alert-message").css("color","#dc3545"):"success"==e&&$(a+" .alert-message").css("color","#28a745"),$(a+" .alert-message").text(i),setTimeout(function(){$(a+" .alert-message").text(""),null!=n&&n()},t)}Office.initialize=function(a){$(document).ready(function(){s=Office.context.roamingSettings,r=Office.context.mailbox.item,$(".content-sender-display-name").text(Office.context.mailbox.item.from.displayName),s.get("user_email")?(e=s.get("user_email"),i=s.get("user_password"),l(t=s.get("user_client_id"),e,i,n="https://"+t+".salesseek.net/api")):($(".page-div").addClass("hidden"),$("#div-authentication-details").removeClass("hidden"))})},$("#btn-login").click(function(a){(function(){var e=!0;$.each($("#div-authentication-details input"),function(i,t){""==$.trim($(t).val())?($(t).siblings(".error-message").html($(t).data("name")+" is required."),e=!1):$(t).siblings(".error-message").val("")});var i=$.trim($("#field-email").val());""==i||i.match(o)||($("#field-email").siblings(".error-message").val("Enter valid email address."),e=!1);return e})()&&(t=$("#field-account-name").val(),e=$("#field-email").val(),i=$("#field-password").val(),n=n="https://"+t+".salesseek.net/api",l(t,e,i,n))}),$("#field-check-organization").change(function(e){$(".field-input-organization").val(""),$(this).is(":checked")?($(".field-input-organization").attr("disabled",!1).focus(),$("#field-select-organization").val("").attr("disabled","disabled"),$("#field-select-organization").siblings(".error-message").html("")):($(".field-input-organization").attr("disabled","disabled"),$("#field-input-organization").siblings(".error-message").html(""),$("#field-select-organization").val("").attr("disabled",!1).focus())}),$(".checkbox").click(function(e){$("#field-check-organization").prop("checked",!$("#field-check-organization").is(":checked")),$("#field-check-organization").trigger("change")}),$(".field-name").change(function(e){m()}),$(".field-auth").change(function(e){if(""==$.trim($(this).val())?($(this).siblings(".error-message").html($(this).data("name")+" is required."),validated=!1):$(this).siblings(".error-message").val(""),"field-email"==$(this).attr("id")){var i=$.trim($(this).val());""==i||i.match(o)||($(this).siblings(".error-message").html("Enter valid email address."),validated=!1)}}),$("#field-account-name").keyup(function(e){$("#help-field-account-name").text($(this).val())}),$(".field-organization").change(function(e){f()}),$("#btn-save-individual").click(function(e){if(m()&&f())if($("#loading-gif-add-individual").removeClass("hidden"),$("#loading-gif-alt").addClass("hidden"),$("#div-add-user-to-crm .detail-section").attr("disabled","disabled"),$("#field-check-organization").is(":checked")){var i=JSON.stringify({name:$("#field-input-organization").val()});$.ajax({type:"POST",url:n+"/organizations",crossDomain:!0,xhrFields:{withCredentials:!0},data:i,dataType:"json",success:function(e,i,t){g(e.id)},error:function(){$("#loading-gif-add-individual").addClass("hidden"),$("#loading-gif-alt").removeClass("hidden"),$("#div-add-user-to-crm .detail-section").attr("disabled",!1),v("error","Could not add "+r.from.displayName+" to SalesSleek. Please try again or contact system administrator.",5e3,"#div-add-user-to-crm")}})}else g($("#field-select-organization").val())})}()}});