package com.qnb.message.controller;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;

import com.qnb.message.entity.User;

@Controller
@RequestMapping(value = "/")
public class UserController {

//	@Autowired
//	public UserRepository countryRepo;

	@Autowired
	public Util myUtil;

	@GetMapping()
	public String showPage() {
		return "index";
	}

	@GetMapping("/success")
	public String showSuccess() {
		return "success";
	}

	@GetMapping("/errors")
	public String showError() {
		return "error";
	}
	
	@RequestMapping(value = "/list", method = RequestMethod.GET)
	public String showDefaultPage(Model model) {
		model.addAttribute("data", myUtil.getAllMessages());
		return "list";
	}

	@PostMapping(value = "/save")
	public String save(User user) throws EncryptedDocumentException, InvalidFormatException, IOException {
		if (user.getName().isEmpty() || user.getSicil().isEmpty() || user.getMessage().isEmpty()) {
			return "redirect:/errors";
		}

		return myUtil.createXLSFile(user);
	}

}
