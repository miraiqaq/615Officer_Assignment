
library(officer)
library(magrittr)
library(tidyverse)
library(readxl)

##  pic a template you like
pres1 <- read_pptx("1.pptx") 

##  get the layout
layout_summary(pres1)
master <- "Office Theme"

#
layout_properties(x = pres1, layout = "Title Slide", master = master )
pres1 %<>%  add_slide(layout = "Title Slide", master = master) %>% 
  ph_with_text(type = "ctrTitle", str = "Advantages of a Bear Market") %>% 
  ph_with_text(type = "subTitle", str = "Yes there is a positive side to a Bear Market!")

#
# layout_properties(x = pres1, layout = "Two Content", master = master )
# pres1 %>%  add_slide(layout = "Two Content", master = master) %>% 
#   ph_with_text(type = "title", str = "Investing in Stocks") %>% 
#   ph_with_text(type = "body", str = "1. Represents ownership in a firm
# 2. Earn a return in two ways
#   * Price of the stock rises over time
#   * Dividends are paid to the stockholder
# 3. Stockholders have claim on all assets", index = 1) %>% 
#   ph_with_text(type = "body", str = "4. Right to vote for directors and on certain issues
# 5. Two types
# - Common stock
#   * Right to vote
#   * Receive dividends
# - Preferred stock
#   * Receive a fixed dividend
#   * Do not usually vote", index = 2)

#
pres1 %<>%  add_slide(layout = "Two Content", master = master  ) %>% 
  ph_with_text(type = "title", str = "Investing in Stock") %>% 
  ph_with_ul(type = "body", index = 1, 
             str_list = c("1. Represents ownership in a firm", 
                          "2. Earn a return in two ways", 
                          "Price of the stock rises over time",
                          "Dividends are paid to the stockholder", 
                          "3. Stockholders have claim on all assets"),
             level_list = c(1,1,2,2,1)) %>% 
  ph_with_ul(type = "body", index = 2, 
             str_list = c("4. Right to vote for directors and on certain issues",
                          "5. Two types",
                          "Common stock",
                          "Right to vote",
                          "Receive dividends",
                          "Preferred stock",
                          "Receive a fixed dividend",
                          "Do not usually vote"),
             level_list =  c(1,1,2,3,3,2,3,3))

#
layout_properties(x = pres1, layout = "Title and Content", master = master )
pres1 %<>% add_slide(layout = "Title and Content", master = master ) %>% 
  ph_with_text(type = "title", str = "Investing in Stocks:Sample Corporate Stock Certificate") %>%
  ph_with_img_at(src = "Picture1.png", height = 5,width = 7,left=4,top=2)
#
layout_properties(x = pres1, layout = "Title and Content", master = master )
pres1 %<>% add_slide(layout = "Title and Content", master = master ) %>% 
  ph_with_text(type = "title", str = "What is a Bear Market?") %>%
  ph_with_text(type = "body", str = "A decline of 15-20% of the broad market coupled with pessimistic sentiment underlying the stock market.",index =1) %>%
  ph_with_img_at(src = "Picture2.png", height = 3,width = 5,left=4,top=4)

#
layout_properties(x = pres1, layout = "Title and Content", master = master )
pres1 %<>% add_slide(layout = "Title and Content", master = master ) %>% 
  ph_with_text(type = "title", str = "Stock Market Indexes: the Dow Jones Industrial Average") %>%
  ph_with_img_at(src = "Picture3.png", height = 4,width = 6,left=4,top=2)

#
layout_properties(x = pres1, layout = "Title and Content", master = master )
pres1 %<>% add_slide(layout = "Title and Content", master = master ) %>% 
  ph_with_text(type = "title", str = "Dow Jones") %>%
  ph_with_img_at(src = "Picture4.png", height = 5,width = 7,left=4,top=2)

#
layout_properties(x = pres1, layout = "Title and Content", master = master )
pres1 %<>% add_slide(layout = "Title and Content", master = master ) %>% 
  ph_with_text(type = "title", str = "The last Bear market") %>%
  ph_with_text(type = "body", str = "Sept. 30, 2002  Dow  7528
Jan. 5, 2004      Dow  10,568
Oct. 8, 2007      Dow   14093",index =1)

#
layout_properties(x = pres1, layout = "Title and Content", master = master )
pres1 %<>% add_slide(layout = "Title and Content", master = master ) %>% 
  ph_with_text(type = "title", str = "What do I do in a Bear Market
") %>%
  ph_with_text(type = "body", str = "Decide whether this is a market correction or the start of something more
Review the stocks you own
Review stocks you wanted to own but were too expensive at time of research
Check your portfolio for balance or the type of stocks you own",index =1)

#

gg_plot <- ggplot(data = iris ) +
  geom_point(mapping = aes(Sepal.Length, Petal.Length), size = 3) +
  theme_minimal()
layout_properties(x = pres1, layout = "Title and Content", master = master )
pres1 %>%  add_slide(layout = "Title and Content", master = master) %>% 
  ph_with_text(type="title", str="ggplot") %>% 
  ph_with_gg(value = gg_plot )



print(pres1, target = "officer assignment.pptx") 

